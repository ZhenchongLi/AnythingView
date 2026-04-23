use gtk::prelude::*;
use webkit::prelude::*;
use std::cell::RefCell;
use std::path::{Path, PathBuf};

use pulldown_cmark::{html as cmark_html, Options, Parser};

use crate::renderer::Renderer;

const HIGHLIGHT_JS: &str = include_str!("../resources/highlight.min.js");
const MERMAID_JS: &str = include_str!("../resources/mermaid.min.js");
const JSZIP_JS: &str = include_str!("../resources/jszip.min.js");
const DOCX_PREVIEW_JS: &str = include_str!("../resources/docx-preview.js");

pub struct WebRenderer {
    webview: webkit::WebView,
    current_path: RefCell<Option<PathBuf>>,
}

impl WebRenderer {
    pub const fn extensions() -> &'static [&'static str] {
        &[
            // Word docs (rendered via docmod CLI)
            "docx", "docmod", "doct",
            // HTML
            "html", "htm",
            // Markdown
            "md", "markdown",
            // LaTeX (compiled via tectonic)
            "tex",
            // LaTeX auxiliary (syntax highlight only)
            "sty", "cls", "bib", "bbl",
            // Subtitles
            "srt", "vtt", "ass", "ssa", "sub", "sbv",
            // Video (native WebKit)
            "mp4", "mov", "m4v", "webm", "m2ts", "ts", "3gp",
            // Video (transcoded via ffmpeg)
            "mkv", "avi", "flv", "wmv", "ogv", "rmvb", "rm", "asf", "vob", "divx", "f4v",
            // Code — languages
            "swift", "cs", "py", "js", "ts", "tsx", "jsx", "go", "rs", "rb", "java",
            "kt", "scala", "c", "h", "cpp", "hpp", "m", "mm", "lua", "r", "pl",
            "php", "dart", "zig", "nim", "ex", "exs", "erl", "hs", "ml", "fs",
            "v", "sv", "vhdl", "asm", "s", "sql",
            // Shell
            "sh", "bash", "zsh", "fish", "bat", "ps1", "cmd",
            // Data / config
            "xml", "json", "yaml", "yml", "toml", "ini", "cfg", "conf",
            "csv", "tsv", "plist", "graphql", "proto",
            // Web / styles
            "css", "scss", "sass", "less",
            // Docs / text
            "rst", "txt", "log", "diff", "patch",
            // Config files
            "env", "editorconfig", "gitignore", "gitattributes", "dockerignore",
            "makefile", "cmake", "gradle",
            // Other
            "lock", "sum", "mod",
        ]
    }

    pub fn supports(ext: &str) -> bool {
        Self::extensions().contains(&ext)
    }

    pub fn new() -> Self {
        let webview = webkit::WebView::new();
        Self {
            webview,
            current_path: RefCell::new(None),
        }
    }

    fn ext_lower(path: &Path) -> String {
        path.extension()
            .and_then(|e| e.to_str())
            .map(|s| s.to_ascii_lowercase())
            .unwrap_or_default()
    }

    fn file_uri(path: &Path) -> String {
        glib::filename_to_uri(path, None)
            .map(|g| g.to_string())
            .unwrap_or_else(|_| format!("file://{}", path.to_string_lossy()))
    }

    fn lang_for(ext: &str) -> String {
        match ext {
            "rs" => "rust".into(),
            "py" => "python".into(),
            "js" | "jsx" => "javascript".into(),
            "ts" | "tsx" => "typescript".into(),
            "sh" | "bash" | "zsh" => "bash".into(),
            "yml" | "yaml" => "yaml".into(),
            "md" | "markdown" => "markdown".into(),
            "h" => "c".into(),
            "hpp" => "cpp".into(),
            "m" | "mm" => "objectivec".into(),
            "kt" => "kotlin".into(),
            "cs" => "csharp".into(),
            "rb" => "ruby".into(),
            "pl" => "perl".into(),
            "ex" | "exs" => "elixir".into(),
            "erl" => "erlang".into(),
            "hs" => "haskell".into(),
            "ml" => "ocaml".into(),
            "fs" => "fsharp".into(),
            "asm" | "s" => "x86asm".into(),
            "fish" => "shell".into(),
            "bat" | "cmd" => "dos".into(),
            "ps1" => "powershell".into(),
            "toml" | "cfg" | "conf" => "ini".into(),
            "plist" | "svg" => "xml".into(),
            "sass" => "scss".into(),
            "rst" | "txt" | "log" => "plaintext".into(),
            "patch" => "diff".into(),
            "proto" => "protobuf".into(),
            "tex" | "sty" | "cls" => "latex".into(),
            "bib" | "bbl" => "tex".into(),
            other => other.to_string(),
        }
    }

    fn load_video_file(&self, path: &Path) {
        let video_uri = Self::file_uri(path);
        let filename = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
        let filename_esc = html_escape::encode_text(filename).to_string();

        let html = format!(
            r#"<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:#000;display:flex;flex-direction:column;height:100vh;overflow:hidden;}}
video{{flex:1;width:100%;min-height:0;background:#000;display:block;}}
#toolbar{{flex:0 0 auto;display:flex;align-items:center;gap:8px;padding:6px 12px;
         background:#111;font-family:system-ui,sans-serif;font-size:12px;color:#aaa;}}
#toolbar button{{padding:4px 10px;border:1px solid rgba(255,255,255,0.15);border-radius:5px;
               background:rgba(255,255,255,0.08);color:#ccc;font-size:12px;cursor:pointer;}}
#toolbar button:hover{{background:rgba(255,255,255,0.15);}}
#sub-label{{color:#666;font-style:italic;}}
#err{{color:#f87171;font-size:12px;display:none;padding:4px 8px;}}
</style></head><body>
<video id="v" controls preload="metadata">
  <source src="{video_uri}">
  <track id="sub-track" kind="subtitles" label="字幕" srclang="und" default>
</video>
<div id="toolbar">
  <button onclick="document.getElementById('f').click()">加载字幕…</button>
  <span id="sub-label">未加载字幕</span>
  <span id="err"></span>
  <input type="file" id="f" accept=".srt,.vtt,.ass,.ssa,.sub,.sbv" style="display:none">
</div>
<script>
function srtToVtt(s){{return 'WEBVTT\n\n'+s.replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/(\d{{2}}:\d{{2}}:\d{{2}}),(\d{{3}})/g,'$1.$2').trim();}}
function assToVtt(s){{var vtt='WEBVTT\n\n',idx=1;s.split('\n').forEach(function(l){{var m=l.match(/^Dialogue:\s*\d+,([\d:.]+),([\d:.]+),[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,(.*)/);if(!m)return;function t(ts){{var p=ts.split(':');return(p[0].length<2?'0'+p[0]:p[0])+':'+p[1]+':'+p[2].replace('.','.'); }}var tx=m[3].replace(/\{{[^}}]*\}}/g,'').replace(/<[^>]+>/g,'');vtt+=(idx++)+'\n'+t(m[1])+' --> '+t(m[2])+'\n'+tx+'\n\n';}});return vtt;}}
document.getElementById('f').onchange=function(e){{
  var file=e.target.files[0];if(!file)return;
  var err=document.getElementById('err');err.style.display='none';
  var reader=new FileReader();
  reader.onerror=function(){{err.textContent='读取失败';err.style.display='inline';}};
  reader.onload=function(ev){{
    var content=ev.target.result;
    var ext=file.name.split('.').pop().toLowerCase();
    var vtt;
    try{{if(ext==='vtt')vtt=content;else if(ext==='ass'||ext==='ssa')vtt=assToVtt(content);else vtt=srtToVtt(content);}}
    catch(ex){{err.textContent='解析失败: '+ex.message;err.style.display='inline';return;}}
    var blob=new Blob([vtt],{{type:'text/vtt'}});
    var url=URL.createObjectURL(blob);
    var track=document.getElementById('sub-track');
    var old=track.src;track.src=url;
    if(old&&old.startsWith('blob:'))URL.revokeObjectURL(old);
    var v=document.getElementById('v');
    for(var i=0;i<v.textTracks.length;i++)v.textTracks[i].mode='showing';
    document.getElementById('sub-label').textContent=file.name;
  }};
  reader.readAsText(file,'UTF-8');
}};
document.getElementById('v').onerror=function(){{
  document.getElementById('err').textContent='无法播放此格式（{filename}）— 取决于系统 GStreamer 插件';
  document.getElementById('err').style.display='inline';
}};
</script></body></html>"#,
            video_uri = video_uri,
            filename = filename_esc,
        );

        self.webview.load_html(&html, Some(&video_uri));
    }

    fn ffmpeg_path() -> Option<String> {
        let candidates = [
            "/usr/bin/ffmpeg",
            "/usr/local/bin/ffmpeg",
            "/opt/homebrew/bin/ffmpeg",
        ];
        for c in &candidates {
            if std::path::Path::new(c).exists() {
                return Some(c.to_string());
            }
        }
        // Check HOME/.local/bin/ffmpeg
        if let Ok(home) = std::env::var("HOME") {
            let p = format!("{}/.local/bin/ffmpeg", home);
            if std::path::Path::new(&p).exists() {
                return Some(p);
            }
        }
        None
    }

    fn transcode_and_play(&self, path: &Path) {
        let filename = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
        let filename_esc = html_escape::encode_text(filename).to_string();

        let ffmpeg = match Self::ffmpeg_path() {
            Some(p) => p,
            None => {
                let html = format!(
                    r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
body{{background:#000;color:#aaa;font-family:system-ui,sans-serif;
     display:flex;flex-direction:column;justify-content:center;align-items:center;
     height:100vh;gap:12px;margin:0;}}
code{{color:#f87171;font-size:13px;}}
</style></head><body>
<div>无法播放 <b>{name}</b></div>
<div>需要 ffmpeg：<code>sudo apt install ffmpeg</code></div>
</body></html>"#,
                    name = filename_esc
                );
                self.webview.load_html(&html, None);
                return;
            }
        };

        // Show transcoding message (may not render before blocking, but consistent with tex pattern)
        let loading_html = format!(
            r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
body{{background:#000;color:#aaa;font-family:system-ui,sans-serif;
     display:flex;flex-direction:column;justify-content:center;align-items:center;
     height:100vh;gap:12px;margin:0;}}
</style></head><body>
<div>正在转码 <b>{name}</b>…</div>
</body></html>"#,
            name = filename_esc
        );
        self.webview.load_html(&loading_html, None);

        let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("video");
        let out_path = std::env::temp_dir().join(format!("anyview-transcode-{}.mp4", stem));

        let result = std::process::Command::new(&ffmpeg)
            .args([
                "-y", "-i", path.to_str().unwrap_or(""),
                "-c:v", "copy", "-c:a", "aac", "-movflags", "faststart",
                out_path.to_str().unwrap_or(""),
            ])
            .output();

        match result {
            Ok(_) if out_path.exists() => {
                self.load_video_file(&out_path);
            }
            Ok(out) => {
                let stderr = String::from_utf8_lossy(&out.stderr);
                let html = format!(
                    r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
body{{background:#000;color:#aaa;font-family:system-ui,sans-serif;padding:24px;margin:0;}}
pre{{color:#f87171;font-size:11px;white-space:pre-wrap;margin-top:12px;}}
</style></head><body>
<div>转码失败：{name}</div>
<pre>{err}</pre>
</body></html>"#,
                    name = filename_esc,
                    err = html_escape::encode_text(&stderr)
                );
                self.webview.load_html(&html, None);
            }
            Err(e) => {
                self.show_error(&format!("ffmpeg 启动失败: {}", e));
            }
        }
    }

    fn load_subtitle_file(&self, path: &Path) {
        let raw = std::fs::read(path)
            .ok()
            .and_then(|b| String::from_utf8(b.clone()).ok()
                .or_else(|| String::from_utf8_lossy(&b).to_string().into()))
            .unwrap_or_default();

        let source_escaped = html_escape::encode_text(&raw).to_string();
        let ext = Self::ext_lower(path);

        // Escape for JS template literal
        let js_source = raw
            .replace('\\', "\\\\")
            .replace('`', "\\`")
            .replace('$', "\\$");

        let html = format!(
            r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light dark">
<style>
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{font-family:system-ui,-apple-system,"Segoe UI",sans-serif;font-size:14px;background:#fff;color:#1a1a1a;}}
#preview{{padding:0 0 40px;}}
.entry{{display:flex;gap:0;border-bottom:1px solid #f0f0f0;}}
.entry:hover{{background:#f9fafb;}}
.num{{flex:0 0 48px;padding:10px 8px 10px 16px;color:#aaa;font-size:12px;text-align:right;user-select:none;}}
.time{{flex:0 0 220px;padding:10px 12px;font-family:"JetBrains Mono","Fira Code",Menlo,monospace;font-size:11px;color:#888;white-space:nowrap;}}
.text{{flex:1;padding:10px 16px 10px 0;line-height:1.5;}}
.header{{position:sticky;top:0;z-index:10;padding:8px 16px;font-size:12px;color:#888;
         background:#fff;border-bottom:1px solid #e5e7eb;display:flex;gap:12px;}}
.header .th-num{{flex:0 0 48px;text-align:right;}}
.header .th-time{{flex:0 0 220px;padding-left:12px;}}
.header .th-text{{flex:1;}}
#source{{display:none;margin:0;padding:20px 24px;white-space:pre-wrap;word-wrap:break-word;
        font-family:"JetBrains Mono","Fira Code",Menlo,monospace;font-size:13px;line-height:1.5;}}
.toggle-btn{{position:fixed;top:12px;right:16px;z-index:9999;padding:4px 12px;
            border:1px solid rgba(0,0,0,0.15);border-radius:6px;
            background:rgba(255,255,255,0.9);color:#333;
            font-size:12px;font-family:system-ui,sans-serif;cursor:pointer;}}
@media(prefers-color-scheme:dark){{
body{{background:#1a1a1a;color:#d4d4d4;}}
.entry{{border-bottom-color:#2a2a2a;}}
.entry:hover{{background:#222;}}
.header{{background:#1a1a1a;border-bottom-color:#333;}}
.toggle-btn{{background:rgba(40,40,40,0.9);border-color:rgba(255,255,255,0.15);color:#ccc;}}
}}
</style></head><body>
<button class="toggle-btn" onclick="toggle()">&lt;/&gt; Source</button>
<div id="preview">
  <div class="header"><span class="th-num">#</span><span class="th-time">Timecode</span><span class="th-text">Text</span></div>
  <div id="entries"></div>
</div>
<pre id="source">{source}</pre>
<script>
var showing='preview';
function toggle(){{
var p=document.getElementById('preview');var s=document.getElementById('source');var btn=document.querySelector('.toggle-btn');
if(showing==='preview'){{p.style.display='none';s.style.display='block';btn.textContent='Preview';showing='source';}}
else{{s.style.display='none';p.style.display='block';btn.innerHTML='&lt;/&gt; Source';showing='preview';}}
}}
function esc(s){{return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}}
function inlineStyle(s){{return s.replace(/<i>(.*?)<\/i>/g,'<em>$1</em>').replace(/<b>(.*?)<\/b>/g,'<b>$1</b>').replace(/<[^>]+>/g,'');}}
var raw=`{js_source}`;
var ext='{ext}';
var entries=[];
if(ext==='vtt'){{
  raw.replace(/\r\n/g,'\n').split(/\n\s*\n/).forEach(function(b){{
    b=b.trim();if(!b||b==='WEBVTT')return;
    var lines=b.split('\n');
    var ti=lines.findIndex(function(l){{return l.indexOf('-->')!==-1;}});
    if(ti===-1)return;
    var num=ti>0?lines[0]:String(entries.length+1);
    entries.push({{num:num,time:lines[ti],text:lines.slice(ti+1).join('<br>')}});
  }});
}}else{{
  raw.replace(/\r\n/g,'\n').split(/\n\s*\n/).forEach(function(b){{
    b=b.trim();if(!b)return;
    var lines=b.split('\n');if(lines.length<2)return;
    var time=lines[1].trim();if(time.indexOf('-->')===-1)return;
    entries.push({{num:lines[0].trim(),time:time,text:lines.slice(2).join('<br>')}});
  }});
}}
var c=document.getElementById('entries');
c.innerHTML=entries.length===0
  ?'<div style="padding:24px 16px;color:#888;font-size:13px;">No entries found — see source view.</div>'
  :entries.map(function(e){{return '<div class="entry"><div class="num">'+esc(e.num)+'</div><div class="time">'+esc(e.time)+'</div><div class="text">'+inlineStyle(e.text)+'</div></div>';}}).join('');
</script></body></html>"#,
            source = source_escaped,
            js_source = js_source,
            ext = ext,
        );

        self.webview.load_html(&html, None);
    }

    fn tectonic_path() -> Option<String> {
        let candidates = [
            "/usr/local/bin/tectonic",
            "/usr/bin/tectonic",
        ];
        for c in &candidates {
            if std::fs::metadata(c).map(|m| m.is_file()).unwrap_or(false) {
                return Some(c.to_string());
            }
        }
        if let Ok(home) = std::env::var("HOME") {
            let local = format!("{}/.local/bin/tectonic", home);
            if std::fs::metadata(&local).map(|m| m.is_file()).unwrap_or(false) {
                return Some(local);
            }
        }
        None
    }

    fn tex_source_html(escaped_source: &str, status_msg: &str, is_error: bool) -> String {
        let status_color = if is_error { "#b91c1c" } else { "#888" };
        let status_html = if status_msg.is_empty() {
            String::new()
        } else {
            format!(
                r#"<div style="position:fixed;top:12px;right:16px;z-index:9999;padding:6px 12px;font:12px system-ui,sans-serif;color:{};background:rgba(0,0,0,0.06);border-radius:4px;max-width:60%;white-space:pre-wrap;">{}</div>"#,
                status_color,
                html_escape::encode_text(status_msg)
            )
        };
        format!(
            r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light dark">
<style>
body{{margin:0;padding:20px 24px;font-family:"JetBrains Mono","Fira Code",Menlo,monospace;font-size:13px;background:#f8f9fa;color:#1a1a1a;}}
pre{{margin:0;line-height:1.5;white-space:pre-wrap;word-wrap:break-word;tab-size:4;}}
.hljs{{display:block;overflow-x:auto;padding:0;color:#333;background:transparent;}}
.hljs-comment,.hljs-quote{{color:#998;font-style:italic;}}
.hljs-keyword,.hljs-selector-tag{{color:#333;font-weight:bold;}}
.hljs-string,.hljs-doctag{{color:#d14;}}
.hljs-title,.hljs-section{{color:#900;font-weight:bold;}}
.hljs-built_in{{color:#0086b3;}}
@media(prefers-color-scheme:dark){{
body{{background:#1e1e1e;color:#d4d4d4;}}
.hljs{{color:#abb2bf;}}
.hljs-keyword{{color:#c678dd;}}
.hljs-string{{color:#98c379;}}
.hljs-built_in{{color:#e6c07b;}}
}}
</style>
<script>{hljs}</script>
</head><body>
{status}
<pre><code class="language-latex">{source}</code></pre>
<script>if(window.hljs){{hljs.highlightAll();}}</script>
</body></html>"#,
            hljs = HIGHLIGHT_JS,
            status = status_html,
            source = escaped_source,
        )
    }

    fn tex_pdf_html(escaped_source: &str, pdf_uri: &str) -> String {
        format!(
            r#"<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light dark">
<style>
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:#fff;}}
iframe{{width:100%;height:100vh;border:none;display:block;}}
#source{{display:none;margin:0;padding:20px 24px;white-space:pre-wrap;word-wrap:break-word;
        font-family:"JetBrains Mono","Fira Code",Menlo,monospace;font-size:13px;line-height:1.5;
        tab-size:4;color:#1a1a1a;background:#f8f9fa;min-height:100vh;}}
.toggle-btn{{position:fixed;top:12px;right:16px;z-index:9999;padding:4px 12px;
            border:1px solid rgba(0,0,0,0.15);border-radius:6px;
            background:rgba(255,255,255,0.9);color:#333;
            font-size:12px;font-family:system-ui,sans-serif;cursor:pointer;
            backdrop-filter:blur(8px);}}
.toggle-btn:hover{{background:rgba(240,240,240,0.95);}}
.hljs{{display:block;overflow-x:auto;padding:0;color:#333;background:transparent;}}
.hljs-comment,.hljs-quote{{color:#998;font-style:italic;}}
.hljs-keyword,.hljs-selector-tag{{color:#333;font-weight:bold;}}
.hljs-string,.hljs-doctag{{color:#d14;}}
.hljs-title,.hljs-section{{color:#900;font-weight:bold;}}
.hljs-built_in{{color:#0086b3;}}
@media(prefers-color-scheme:dark){{
body{{background:#1a1a1a;}}
#source{{color:#d4d4d4;background:#1e1e1e;}}
.toggle-btn{{background:rgba(40,40,40,0.9);border-color:rgba(255,255,255,0.15);color:#ccc;}}
.hljs{{color:#abb2bf;}}
.hljs-keyword{{color:#c678dd;}}
.hljs-string{{color:#98c379;}}
.hljs-built_in{{color:#e6c07b;}}
}}
</style>
<script>{hljs}</script>
</head><body>
<button class="toggle-btn" onclick="toggle()">&lt;/&gt; Source</button>
<iframe id="preview" src="{pdf}"></iframe>
<pre id="source"><code class="language-latex">{source}</code></pre>
<script>
var showing='preview';
function toggle(){{
var p=document.getElementById('preview');
var s=document.getElementById('source');
var btn=document.querySelector('.toggle-btn');
if(showing==='preview'){{p.style.display='none';s.style.display='block';btn.textContent='PDF';showing='source';if(window.hljs)hljs.highlightAll();}}
else{{s.style.display='none';p.style.display='block';btn.innerHTML='&lt;/&gt; Source';showing='preview';}}
}}
</script>
</body></html>"#,
            hljs = HIGHLIGHT_JS,
            pdf = pdf_uri,
            source = escaped_source,
        )
    }

    fn load_tex_file(&self, path: &Path) {
        let source = match std::fs::read_to_string(path) {
            Ok(s) => s,
            Err(e) => { self.show_error(&format!("Failed to read file: {}", e)); return; }
        };
        let escaped = html_escape::encode_text(&source).to_string();

        // Subfile check — no \begin{document} means it's meant to be \input'd
        if !source.contains(r"\begin{document}") {
            let html = Self::tex_source_html(&escaped, "Subfile (no \\begin{document}) — source only", false);
            self.webview.load_html(&html, None);
            return;
        }

        let tectonic = match Self::tectonic_path() {
            Some(p) => p,
            None => {
                let html = Self::tex_source_html(
                    &escaped,
                    "tectonic not found — install: cargo install tectonic  or  apt install tectonic",
                    true,
                );
                self.webview.load_html(&html, None);
                return;
            }
        };

        // Show source + "Compiling…" while tectonic runs
        self.webview.load_html(&Self::tex_source_html(&escaped, "Compiling\u{2026}", false), None);

        let unique = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        let tmp_dir = std::env::temp_dir().join(format!("anyview-tex-{}", unique));
        let _ = std::fs::create_dir_all(&tmp_dir);

        let output = std::process::Command::new(&tectonic)
            .args(["-Z", "continue-on-errors", "--outdir", tmp_dir.to_str().unwrap_or("")])
            .arg(path)
            .output();

        let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("output");
        let pdf_path = tmp_dir.join(format!("{}.pdf", stem));

        if pdf_path.exists() {
            let pdf_uri = format!("file://{}", pdf_path.to_string_lossy());
            let html = Self::tex_pdf_html(&escaped, &pdf_uri);
            self.webview.load_html(&html, None);
        } else {
            let err_msg = output
                .ok()
                .and_then(|o| String::from_utf8(o.stderr).ok())
                .unwrap_or_else(|| "Compilation failed".to_string());
            let html = Self::tex_source_html(&escaped, err_msg.trim(), true);
            self.webview.load_html(&html, None);
        }
    }

    fn load_html_file(&self, path: &Path) {
        // Preferred: let WebKit handle relative paths natively via file URI.
        let uri = Self::file_uri(path);
        self.webview.load_uri(&uri);
    }

    fn load_docx_file(&self, path: &Path) {
        // Render .docx natively in the browser via docx-preview (docxjs):
        // read the raw bytes, base64-encode them, and let the embedded
        // script rehydrate the bytes into a Blob + renderAsync.
        let bytes = match std::fs::read(path) {
            Ok(b) => b,
            Err(e) => {
                self.show_error(&format!("Failed to read file: {}", e));
                return;
            }
        };

        use base64::{engine::general_purpose::STANDARD, Engine as _};
        let b64 = STANDARD.encode(&bytes);

        let html = format!(
            r#"<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="color-scheme" content="light dark">
<style>
  :root {{ color-scheme: light dark; }}
  html, body {{ margin: 0; padding: 0; background: #e5e7eb; }}
  #container {{ padding: 20px 0; }}
  #container .docx-wrapper {{ background: #9CA3AF; padding: 30px; display: flex; flex-flow: column; align-items: center; }}
  #container .docx-wrapper > section.docx {{ background: #fff; box-shadow: 0 0 10px rgba(0,0,0,0.5); margin-bottom: 30px; }}
  #container .docx {{ color: #000; }}
  @media (prefers-color-scheme: dark) {{
    html, body {{ background: #1a1a1a; }}
    #container .docx-wrapper {{ background: #2a2a2a; }}
  }}
  .status {{ font-family: system-ui, sans-serif; padding: 40px; color: #333; text-align: center; }}
  .status.err {{ color: #c00; }}
</style>
<script>{jszip}</script>
<script>{docxjs}</script>
</head>
<body>
<div id="container"><div class="status">Rendering…</div></div>
<script>
(function() {{
  var b64 = "{b64}";
  try {{
    var bin = atob(b64);
    var len = bin.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    var blob = new Blob([bytes], {{ type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }});
    var container = document.getElementById("container");
    container.innerHTML = "";
    docx.renderAsync(blob, container, null, {{
      className: "docx",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      experimental: false,
      trimXmlDeclaration: true,
      useBase64URL: true,
      renderChanges: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      debug: false
    }}).catch(function(err) {{
      container.innerHTML = '<div class="status err">Render failed: ' + (err && err.message ? err.message : err) + '</div>';
    }});
  }} catch (err) {{
    document.getElementById("container").innerHTML = '<div class="status err">Decode failed: ' + err.message + '</div>';
  }}
}})();
</script>
</body>
</html>"#,
            jszip = JSZIP_JS,
            docxjs = DOCX_PREVIEW_JS,
            b64 = b64,
        );

        let base_uri = Self::file_uri(path);
        self.webview.load_html(&html, Some(&base_uri));
    }

    fn load_docmod_file(&self, path: &Path) {
        match crate::docmod_cli::render(path) {
            Ok(html) => {
                // docmod produces a full HTML doc. Load directly so its
                // embedded styles/scripts run as-is.
                let base_uri = Self::file_uri(path);
                self.webview.load_html(&html, Some(&base_uri));
            }
            Err(msg) => {
                self.show_error(&msg);
            }
        }
    }

    fn load_markdown_file(&self, path: &Path) {
        let raw = match std::fs::read_to_string(path) {
            Ok(s) => s,
            Err(e) => {
                self.show_error(&format!("Failed to read file: {}", e));
                return;
            }
        };

        let mut opts = Options::empty();
        opts.insert(Options::ENABLE_TABLES);
        opts.insert(Options::ENABLE_FOOTNOTES);
        opts.insert(Options::ENABLE_STRIKETHROUGH);
        opts.insert(Options::ENABLE_TASKLISTS);
        opts.insert(Options::ENABLE_SMART_PUNCTUATION);
        opts.insert(Options::ENABLE_HEADING_ATTRIBUTES);

        let parser = Parser::new_ext(&raw, opts);
        let mut body = String::new();
        cmark_html::push_html(&mut body, parser);

        let wrapped = Self::wrap_document(&body);
        let base_uri = Self::file_uri(path);
        self.webview.load_html(&wrapped, Some(&base_uri));
    }

    fn load_code_file(&self, path: &Path) {
        let raw = match std::fs::read_to_string(path) {
            Ok(s) => s,
            Err(e) => {
                self.show_error(&format!("Failed to read file: {}", e));
                return;
            }
        };
        let ext = Self::ext_lower(path);
        let lang = Self::lang_for(&ext);
        let escaped = html_escape::encode_text(&raw).to_string();
        let filename = path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("")
            .to_string();
        let filename_escaped = html_escape::encode_text(&filename).to_string();
        let line_count = raw.lines().count().max(if raw.is_empty() { 0 } else { 1 });

        let body = format!(
            "<div class=\"file-header\">{} lines &middot; {}</div>\n<pre><code class=\"language-{}\">{}</code></pre>",
            line_count, filename_escaped, lang, escaped
        );

        let wrapped = Self::wrap_document(&body);
        let base_uri = Self::file_uri(path);
        self.webview.load_html(&wrapped, Some(&base_uri));
    }

    fn show_error(&self, message: &str) {
        let escaped = html_escape::encode_text(message).to_string();
        let html = format!(
            r#"<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
body {{ font-family: system-ui, -apple-system, "Segoe UI", sans-serif; padding: 40px; color: #333; background: #fff; }}
h2 {{ color: #c00; }}
pre {{ white-space: pre-wrap; word-wrap: break-word; background: #f5f5f5; padding: 12px; border-radius: 6px; }}
@media (prefers-color-scheme: dark) {{
  body {{ background: #1a1a1a; color: #d4d4d4; }}
  pre {{ background: #252525; }}
}}
</style></head><body>
<h2>Error</h2>
<pre>{}</pre>
</body></html>"#,
            escaped
        );
        self.webview.load_html(&html, None);
    }

    /// Wraps rendered HTML body in the full template (head + scripts + styles).
    fn wrap_document(body_html: &str) -> String {
        format!(
            r#"<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="color-scheme" content="light dark">
<style>
:root {{
  color-scheme: light dark;
}}
* {{ box-sizing: border-box; }}
body {{
  font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  font-size: 15px;
  line-height: 1.6;
  color: #1a1a1a;
  background: #ffffff;
  max-width: 900px;
  margin: 0 auto;
  padding: 24px 32px;
}}
h1, h2, h3, h4, h5, h6 {{ margin: 1em 0 0.5em; line-height: 1.25; }}
h1 {{ font-size: 1.8em; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.3em; }}
h2 {{ font-size: 1.4em; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.3em; }}
h3 {{ font-size: 1.2em; }}
p {{ margin: 0.75em 0; }}
a {{ color: #2563eb; }}
code {{
  font-family: "JetBrains Mono", "Fira Code", "SF Mono", Menlo, Consolas, monospace;
  font-size: 0.9em;
  background: #f3f4f6;
  padding: 2px 5px;
  border-radius: 3px;
}}
pre {{
  background: #f6f8fa;
  border-radius: 6px;
  overflow-x: auto;
  margin: 1em 0;
  padding: 12px 14px;
  line-height: 1.5;
}}
pre code {{
  background: none;
  padding: 0;
  font-size: 0.88em;
  border-radius: 0;
}}
blockquote {{
  margin: 1em 0;
  padding: 0 1em;
  border-left: 4px solid #d1d5db;
  color: #6b7280;
}}
table {{ border-collapse: collapse; width: 100%; margin: 1em 0; }}
th, td {{ border: 1px solid #d1d5db; padding: 8px 12px; text-align: left; }}
th {{ background: #f9fafb; font-weight: 600; }}
img {{ max-width: 100%; }}
hr {{ border: none; border-top: 1px solid #e5e7eb; margin: 2em 0; }}
ul, ol {{ padding-left: 2em; }}
li {{ margin: 0.25em 0; }}
.mermaid {{ display: flex; justify-content: center; margin: 1em 0; }}
.mermaid svg {{ max-width: 100%; height: auto; }}
.file-header {{
  color: #888;
  font-size: 12px;
  margin-bottom: 12px;
  padding-bottom: 8px;
  border-bottom: 1px solid #e5e7eb;
}}
@media (prefers-color-scheme: dark) {{
  body {{ background: #1a1a1a; color: #d4d4d4; }}
  h1, h2 {{ border-bottom-color: #333; }}
  code {{ background: #2d2d2d; }}
  pre {{ background: #1e1e1e; }}
  blockquote {{ border-left-color: #555; color: #999; }}
  th, td {{ border-color: #444; }}
  th {{ background: #252525; }}
  hr {{ border-top-color: #333; }}
  a {{ color: #60a5fa; }}
  .file-header {{ border-bottom-color: #333; color: #888; }}
}}
</style>
<script>{hljs}</script>
<script>{mermaid}</script>
</head>
<body>
{body}
<script>
(function() {{
  // Convert mermaid code blocks to <div class="mermaid"> BEFORE mermaid runs.
  var blocks = document.querySelectorAll('pre > code.language-mermaid');
  for (var i = 0; i < blocks.length; i++) {{
    var code = blocks[i];
    var pre = code.parentNode;
    var div = document.createElement('div');
    div.className = 'mermaid';
    div.textContent = code.textContent;
    pre.parentNode.replaceChild(div, pre);
  }}
  try {{ if (window.hljs) hljs.highlightAll(); }} catch (e) {{}}
  try {{
    if (window.mermaid) {{
      var isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      mermaid.initialize({{ startOnLoad: true, theme: isDark ? 'dark' : 'default' }});
    }}
  }} catch (e) {{}}
}})();
</script>
</body>
</html>
"#,
            hljs = HIGHLIGHT_JS,
            mermaid = MERMAID_JS,
            body = body_html
        )
    }
}

impl Renderer for WebRenderer {
    fn widget(&self) -> gtk::Widget {
        self.webview.clone().upcast()
    }

    fn load(&self, path: &Path) {
        *self.current_path.borrow_mut() = Some(path.to_path_buf());
        let ext = Self::ext_lower(path);
        match ext.as_str() {
            "docx" => self.load_docx_file(path),
            "docmod" | "doct" => self.load_docmod_file(path),
            "html" | "htm" => self.load_html_file(path),
            "md" | "markdown" => self.load_markdown_file(path),
            "tex" => self.load_tex_file(path),
            "srt" | "vtt" | "ass" | "ssa" | "sub" | "sbv" => self.load_subtitle_file(path),
            "mp4" | "mov" | "m4v" | "webm" | "m2ts" | "ts" | "3gp" => self.load_video_file(path),
            "mkv" | "avi" | "flv" | "wmv" | "ogv" | "rmvb" | "rm" | "asf" | "vob"
            | "divx" | "f4v" => self.transcode_and_play(path),
            _ => self.load_code_file(path),
        }
    }

    fn set_zoom(&self, level: f64) {
        self.webview.set_zoom_level(level);
    }
}
