import Cocoa
import PDFKit
import WebKit

private class AnyViewWebView: WKWebView {
    override func performKeyEquivalent(with event: NSEvent) -> Bool {
        // Block WKWebView's native Cmd+R reload (which blanks loadHTMLString content).
        // Send directly to AppDelegate to avoid hitting WKWebView.reload() via responder chain.
        if event.modifierFlags.intersection(.deviceIndependentFlagsMask) == .command,
           event.charactersIgnoringModifiers == "r" {
            NSApp.sendAction(#selector(AppDelegate.reload(_:)), to: NSApp.delegate, from: self)
            return true
        }
        return super.performKeyEquivalent(with: event)
    }
}

/// Renders documents using WKWebView.
/// Handles docx, docmod, doct, html, markdown, and code files.
class WebRenderer: NSObject, ViewerRenderer, SupportsFind, WKNavigationDelegate {
    static let docExtensions: Set<String> = ["docmod", "doct", "docx"]
    static let htmlExtensions: Set<String> = ["html", "htm"]
    static let markdownExtensions: Set<String> = ["md", "markdown"]
    static let texExtensions: Set<String> = ["tex"]
    static let subtitleExtensions: Set<String> = ["srt", "vtt", "ass", "ssa", "sub", "sbv"]
    static let videoExtensions: Set<String> = [
        "mp4", "mov", "m4v", "webm", "m2ts", "ts", "3gp",
        "mkv", "avi", "flv", "wmv", "ogv", "rmvb", "rm", "asf", "vob", "divx", "f4v",
    ]
    private static let nativeVideoExtensions: Set<String> = [
        "mp4", "mov", "m4v", "webm", "m2ts", "ts", "3gp",
    ]
    static let codeExtensions: Set<String> = [
        // Languages
        "swift", "cs", "py", "js", "ts", "tsx", "jsx", "go", "rs", "rb", "java",
        "kt", "scala", "c", "h", "cpp", "hpp", "m", "mm", "lua", "r", "pl",
        "php", "dart", "zig", "nim", "ex", "exs", "erl", "hs", "ml", "fs",
        "v", "sv", "vhdl", "asm", "s", "sql",
        // Shell / scripting
        "sh", "bash", "zsh", "fish", "bat", "ps1", "cmd",
        // Markup / data
        "xml", "json", "yaml", "yml", "toml", "ini", "cfg", "conf",
        "csv", "tsv", "plist", "graphql", "proto",
        // Web
        "css", "scss", "sass", "less", "svg",
        // Docs / text
        "rst", "txt", "log", "diff", "patch",
        // Config
        "env", "editorconfig", "gitignore", "gitattributes", "dockerignore",
        "makefile", "cmake", "gradle", "sln", "csproj", "xcodeproj",
        // Other
        "lock", "sum", "mod",
        // LaTeX auxiliary
        "sty", "cls", "bib", "bbl",
    ]
    static let supportedExtensions: Set<String> = docExtensions
        .union(htmlExtensions)
        .union(markdownExtensions)
        .union(texExtensions)
        .union(subtitleExtensions)
        .union(videoExtensions)
        .union(codeExtensions)

    static let mermaidScript: String = {
        guard let url = Bundle.module.url(forResource: "mermaid.min", withExtension: "js"),
              let content = try? String(contentsOf: url, encoding: .utf8) else {
            return ""
        }
        return content
    }()

    static let highlightScript: String = {
        guard let url = Bundle.module.url(forResource: "highlight.min", withExtension: "js"),
              let content = try? String(contentsOf: url, encoding: .utf8) else {
            return ""
        }
        return content
    }()

    static let docxPreviewScript: String = {
        guard let url = Bundle.module.url(forResource: "docx-preview.min", withExtension: "js"),
              let content = try? String(contentsOf: url, encoding: .utf8) else {
            return ""
        }
        return content
    }()

    static let jszipScript: String = {
        guard let url = Bundle.module.url(forResource: "jszip.min", withExtension: "js"),
              let content = try? String(contentsOf: url, encoding: .utf8) else {
            return ""
        }
        return content
    }()

    static let hljsLatexScript: String = {
        guard let url = Bundle.module.url(forResource: "hljs-latex", withExtension: "js"),
              let content = try? String(contentsOf: url, encoding: .utf8) else {
            return ""
        }
        return content
    }()

    private let containerView: NSView
    private let webView: AnyViewWebView
    private let pdfView: PDFView
    private let texToggleBtn: NSButton
    private let lock = NSLock()
    private var _tempDir: String?
    private var tempDir: String? {
        get { lock.lock(); defer { lock.unlock() }; return _tempDir }
        set { lock.lock(); defer { lock.unlock() }; _tempDir = newValue }
    }
    private var currentFilePath: String?
    private var zoomLevel: CGFloat = 1.0

    // Tex-specific state
    private var texShowingPdf = false
    private var texPdfPath: String?
    private var texSourceHtml: String?
    private var texLastFind: PDFSelection?

    // Video transcode state
    private var transcodeProcess: Process?

    var view: NSView { containerView }

    private var fileExtension: String {
        guard let fp = currentFilePath else { return "" }
        return URL(fileURLWithPath: fp).pathExtension.lowercased()
    }

    override init() {
        containerView = NSView(frame: .zero)

        let config = WKWebViewConfiguration()
        webView = AnyViewWebView(frame: .zero, configuration: config)
        webView.autoresizingMask = [.width, .height]

        pdfView = PDFView(frame: .zero)
        pdfView.autoresizingMask = [.width, .height]
        pdfView.autoScales = true
        pdfView.displayMode = .singlePageContinuous
        pdfView.isHidden = true

        texToggleBtn = NSButton(frame: NSRect(x: 0, y: 0, width: 72, height: 26))
        texToggleBtn.bezelStyle = .rounded
        texToggleBtn.font = NSFont.systemFont(ofSize: 12)
        texToggleBtn.isHidden = true
        texToggleBtn.autoresizingMask = [.minXMargin, .minYMargin]

        containerView.addSubview(webView)
        containerView.addSubview(pdfView)
        containerView.addSubview(texToggleBtn)

        super.init()
        webView.navigationDelegate = self
        texToggleBtn.target = self
        texToggleBtn.action = #selector(toggleTexView)
    }

    deinit {
        if let dir = tempDir {
            ZipExtractor.cleanup(tempDir: dir)
        }
    }

    func load(filePath: String) {
        currentFilePath = filePath

        // Cancel any in-progress transcode
        transcodeProcess?.terminate()
        transcodeProcess = nil

        // Reset tex state
        texShowingPdf = false
        texPdfPath = nil
        texSourceHtml = nil
        texLastFind = nil
        pdfView.document = nil
        pdfView.isHidden = true
        webView.isHidden = false
        texToggleBtn.isHidden = true

        if let dir = tempDir {
            ZipExtractor.cleanup(tempDir: dir)
            tempDir = nil
        }

        let ext = URL(fileURLWithPath: filePath).pathExtension.lowercased()

        if Self.htmlExtensions.contains(ext) {
            loadHtmlFile(filePath)
            return
        }
        if Self.markdownExtensions.contains(ext) {
            loadMarkdownFile(filePath)
            return
        }
        if Self.texExtensions.contains(ext) {
            loadTexFile(filePath)
            return
        }
        if Self.subtitleExtensions.contains(ext) {
            loadSubtitleFile(filePath)
            return
        }
        if Self.videoExtensions.contains(ext) {
            loadVideoFile(filePath)
            return
        }
        if Self.codeExtensions.contains(ext) {
            loadCodeFile(filePath)
            return
        }
        if ext == "docx" {
            loadDocxContent(filePath)
            return
        }

        // .docmod / .doct
        let extractedDir: String
        do {
            extractedDir = try ZipExtractor.extract(zipPath: filePath)
            self.tempDir = extractedDir
        } catch {
            showError("Failed to extract file: \(error.localizedDescription)")
            return
        }

        if ext == "doct" {
            loadDoctContent(filePath, extractedDir: extractedDir)
        } else {
            loadDocmodContent(filePath, extractedDir: extractedDir)
        }
    }

    func setZoom(_ level: CGFloat) {
        zoomLevel = level
        if texShowingPdf {
            pdfView.scaleFactor = level
        } else {
            webView.pageZoom = level
            if Self.htmlExtensions.contains(fileExtension) {
                webView.evaluateJavaScript(
                    "window.__vaSetZoom && window.__vaSetZoom(\(level))",
                    completionHandler: nil
                )
            }
        }
    }

    // MARK: - Find

    func performFind(query: String, forward: Bool, completion: @escaping (Bool) -> Void) {
        if texShowingPdf, let doc = pdfView.document {
            var options: NSString.CompareOptions = [.caseInsensitive]
            if !forward { options.insert(.backwards) }
            let match = doc.findString(query, fromSelection: texLastFind, withOptions: options)
                ?? doc.findString(query, fromSelection: nil, withOptions: options)
            if let match {
                texLastFind = match
                pdfView.setCurrentSelection(match, animate: true)
                pdfView.scrollSelectionToVisible(nil)
                completion(true)
            } else {
                completion(false)
            }
            return
        }
        let config = WKFindConfiguration()
        config.backwards = !forward
        config.wraps = true
        config.caseSensitive = false
        webView.find(query, configuration: config) { result in
            completion(result.matchFound)
        }
    }

    // MARK: - Tex toggle

    @objc private func toggleTexView() {
        if texShowingPdf { showTexSource() } else { showTexPdf() }
    }

    private func showTexPdf() {
        guard let path = texPdfPath,
              let doc = PDFDocument(url: URL(fileURLWithPath: path)) else { return }
        texShowingPdf = true
        texLastFind = nil
        pdfView.document = doc
        webView.isHidden = true
        pdfView.isHidden = false
        texToggleBtn.title = "</>"
        positionTexToggleBtn()
    }

    private func showTexSource() {
        guard let html = texSourceHtml else { return }
        texShowingPdf = false
        pdfView.isHidden = true
        webView.isHidden = false
        texToggleBtn.title = "PDF"
        positionTexToggleBtn()
        webView.loadHTMLString(html, baseURL: nil)
    }

    private func positionTexToggleBtn() {
        let margin: CGFloat = 12
        let w = texToggleBtn.frame.width
        let h = texToggleBtn.frame.height
        texToggleBtn.frame = NSRect(
            x: containerView.bounds.width - w - margin,
            y: containerView.bounds.height - h - margin,
            width: w, height: h
        )
        texToggleBtn.isHidden = false
    }

    // MARK: - WKNavigationDelegate

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        setZoom(zoomLevel)
    }

    func webView(_ webView: WKWebView, decidePolicyFor navigationAction: WKNavigationAction,
                 decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        guard navigationAction.navigationType == .linkActivated,
              let url = navigationAction.request.url else {
            decisionHandler(.allow)
            return
        }
        if let fragment = url.fragment, !fragment.isEmpty,
           let currentURL = webView.url,
           url.scheme == currentURL.scheme,
           url.path == currentURL.path {
            decisionHandler(.allow)
            return
        }
        NSWorkspace.shared.open(url)
        decisionHandler(.cancel)
    }

    // MARK: - Cleanup

    func cleanup() {
        if let dir = tempDir {
            ZipExtractor.cleanup(tempDir: dir)
            tempDir = nil
        }
    }

    // MARK: - HTML File

    private func loadHtmlFile(_ filePath: String) {
        guard let content = try? String(contentsOfFile: filePath, encoding: .utf8) else {
            showError("Failed to read file")
            return
        }

        let sourceEscaped = content
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")

        let srcdocEscaped = content
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "'", with: "&#39;")

        let html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <meta name="color-scheme" content="light dark">
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { background: #fff; }
            iframe { width: 100%; height: 100vh; border: none; display: block; }
            #source { display: none; margin: 0; padding: 20px 24px; white-space: pre-wrap; word-wrap: break-word;
                      font-family: "SF Mono", Menlo, monospace; font-size: 13px; line-height: 1.5;
                      tab-size: 4; color: #1a1a1a; background: #f8f9fa; min-height: 100vh; }
            .toggle-btn { position: fixed; top: 12px; right: 16px; z-index: 9999;
                          width: 32px; height: 32px; border: none; border-radius: 50%;
                          background: rgba(0,0,0,0.06); color: #555;
                          font-size: 14px; font-family: -apple-system, sans-serif; font-weight: 500;
                          cursor: pointer; backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
                          display: flex; align-items: center; justify-content: center;
                          transition: all 0.2s; line-height: 1; }
            .toggle-btn:hover { background: rgba(0,0,0,0.12); transform: scale(1.08); }
            @media (prefers-color-scheme: dark) {
                body { background: #1a1a1a; }
                #source { color: #d4d4d4; background: #1e1e1e; }
                .toggle-btn { background: rgba(255,255,255,0.1); color: #aaa; }
                .toggle-btn:hover { background: rgba(255,255,255,0.18); }
            }
        </style>
        </head>
        <body>
        <button class="toggle-btn" onclick="toggle()">&lt;/&gt;</button>
        <iframe id="preview" srcdoc='\(srcdocEscaped)' onload="if(this.contentDocument&&this.contentDocument.body)this.contentDocument.body.style.zoom=window.__vaZoom"></iframe>
        <pre id="source">\(sourceEscaped)</pre>
        <script>
        window.__vaZoom = 1;
        window.__vaSetZoom = function(z) {
            window.__vaZoom = z;
            var f = document.getElementById('preview');
            if (f && f.contentDocument && f.contentDocument.body) {
                f.contentDocument.body.style.zoom = z;
            }
        };
        var showing = 'preview';
        function toggle() {
            var p = document.getElementById('preview');
            var s = document.getElementById('source');
            var btn = document.querySelector('.toggle-btn');
            if (showing === 'preview') {
                p.style.display = 'none';
                s.style.display = 'block';
                btn.innerHTML = '\u{1f441}';
                showing = 'source';
            } else {
                s.style.display = 'none';
                p.style.display = 'block';
                btn.innerHTML = '&lt;/&gt;';
                showing = 'preview';
            }
        }
        </script>
        </body>
        </html>
        """

        let baseURL = URL(fileURLWithPath: filePath).deletingLastPathComponent()
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: baseURL)
        }
    }

    // MARK: - Markdown File

    private func inlineLocalImages(in markdown: String, baseDir: URL) -> String {
        guard let regex = try? NSRegularExpression(pattern: #"!\[([^\]]*)\]\(([^)#?\s"']+)(?:\s+(?:"[^"]*"|'[^']*'))?\)"#) else {
            return markdown
        }
        var result = markdown
        let matches = regex.matches(in: markdown, range: NSRange(markdown.startIndex..., in: markdown))
        for match in matches.reversed() {
            guard let pathRange = Range(match.range(at: 2), in: result),
                  let altRange  = Range(match.range(at: 1), in: result),
                  let fullRange = Range(match.range,         in: result) else { continue }
            let path = String(result[pathRange])
            if path.hasPrefix("http://") || path.hasPrefix("https://") || path.hasPrefix("data:") { continue }
            let imageURL = baseDir.appendingPathComponent(path)
            guard let data = try? Data(contentsOf: imageURL) else { continue }
            let ext = imageURL.pathExtension.lowercased()
            let mime: String
            switch ext {
            case "jpg", "jpeg": mime = "image/jpeg"
            case "gif":         mime = "image/gif"
            case "svg":         mime = "image/svg+xml"
            case "webp":        mime = "image/webp"
            default:            mime = "image/png"
            }
            let alt = String(result[altRange])
            result.replaceSubrange(fullRange,
                with: "![\(alt)](data:\(mime);base64,\(data.base64EncodedString()))")
        }
        return result
    }

    private func loadMarkdownFile(_ filePath: String) {
        guard let raw = try? String(contentsOfFile: filePath, encoding: .utf8) else {
            showError("Failed to read file")
            return
        }
        let baseDir = URL(fileURLWithPath: filePath).deletingLastPathComponent()
        let content = inlineLocalImages(in: raw, baseDir: baseDir)

        let jsEscaped = content
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "`", with: "\\`")
            .replacingOccurrences(of: "$", with: "\\$")
        let jsEscapedSource = raw
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "`", with: "\\`")
            .replacingOccurrences(of: "$", with: "\\$")

        let mermaidInline = content.contains("```mermaid")
            ? "<script>\(Self.mermaidScript)</script>"
            : ""
        let highlightInline = Self.highlightScript.isEmpty ? "" : "<script>\(Self.highlightScript)</script>"

        let html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <meta name="color-scheme" content="light dark">
        <style>
            body { max-width: 800px; margin: 0 auto; padding: 24px 32px; font-family: -apple-system, sans-serif;
                   font-size: 15px; line-height: 1.6; color: #1a1a1a; background: #fff; }
            h1 { font-size: 1.8em; margin: 1em 0 0.5em; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.3em; }
            h2 { font-size: 1.4em; margin: 1em 0 0.5em; border-bottom: 1px solid #e5e7eb; padding-bottom: 0.3em; }
            h3 { font-size: 1.2em; margin: 1em 0 0.5em; }
            h4, h5, h6 { font-size: 1em; margin: 1em 0 0.5em; }
            code { font-family: "SF Mono", Menlo, monospace; font-size: 0.9em; background: #f3f4f6;
                   padding: 2px 5px; border-radius: 3px; }
            pre { border-radius: 6px; overflow-x: auto; line-height: 1.5; margin: 1em 0; }
            pre code { font-size: 0.88em; border-radius: 0; background: none; padding: 0; }
            .hljs{display:block;overflow-x:auto;padding:1em;color:#333;background:#f8f8f8;}
            .hljs-comment,.hljs-quote{color:#998;font-style:italic;}
            .hljs-keyword,.hljs-selector-tag,.hljs-subst{color:#333;font-weight:bold;}
            .hljs-number,.hljs-literal,.hljs-variable,.hljs-template-variable,.hljs-tag .hljs-attr{color:#008080;}
            .hljs-string,.hljs-doctag{color:#d14;}
            .hljs-title,.hljs-section,.hljs-selector-id{color:#900;font-weight:bold;}
            .hljs-subst{font-weight:normal;}
            .hljs-type,.hljs-class .hljs-title{color:#458;font-weight:bold;}
            .hljs-tag,.hljs-name,.hljs-attribute{color:#000080;font-weight:normal;}
            .hljs-regexp,.hljs-link{color:#009926;}
            .hljs-symbol,.hljs-bullet{color:#990073;}
            .hljs-built_in,.hljs-builtin-name{color:#0086b3;}
            .hljs-meta{color:#999;font-weight:bold;}
            .hljs-deletion{background:#fdd;}.hljs-addition{background:#dfd;}
            .mermaid { display: flex; justify-content: center; margin: 1em 0; }
            .mermaid svg { max-width: 100%; height: auto; }
            blockquote { margin: 1em 0; padding: 0 1em; border-left: 4px solid #d1d5db; color: #6b7280; }
            table { border-collapse: collapse; width: 100%; margin: 1em 0; }
            th, td { border: 1px solid #d1d5db; padding: 8px 12px; text-align: left; }
            th { background: #f9fafb; font-weight: 600; }
            img { max-width: 100%; }
            hr { border: none; border-top: 1px solid #e5e7eb; margin: 2em 0; }
            a { color: #2563eb; }
            ul, ol { padding-left: 2em; }
            li { margin: 0.25em 0; }
            #source { display: none; margin: 0; padding: 0; white-space: pre-wrap; word-wrap: break-word;
                      font-family: "SF Mono", Menlo, monospace; font-size: 13px; line-height: 1.5;
                      tab-size: 4; color: #1a1a1a; background: none; }
            .toggle-btn { position: fixed; top: 12px; right: 16px; z-index: 9999;
                          width: 32px; height: 32px; border: none; border-radius: 50%;
                          background: rgba(0,0,0,0.06); color: #555;
                          font-size: 14px; font-family: -apple-system, sans-serif; font-weight: 500;
                          cursor: pointer; backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
                          display: flex; align-items: center; justify-content: center;
                          transition: all 0.2s; line-height: 1; }
            .toggle-btn:hover { background: rgba(0,0,0,0.12); transform: scale(1.08); }
            @media (prefers-color-scheme: dark) {
                body { background: #1a1a1a; color: #d4d4d4; }
                h1, h2 { border-bottom-color: #333; }
                code { background: #2d2d2d; }
                blockquote { border-left-color: #555; color: #999; }
                th, td { border-color: #444; }
                th { background: #252525; }
                hr { border-top-color: #333; }
                a { color: #60a5fa; }
                #source { color: #d4d4d4; }
                .toggle-btn { background: rgba(255,255,255,0.1); color: #aaa; }
                .toggle-btn:hover { background: rgba(255,255,255,0.18); }
                .hljs{color:#abb2bf;background:#282c34;}
                .hljs-comment,.hljs-quote{color:#5c6370;font-style:italic;}
                .hljs-doctag,.hljs-keyword,.hljs-formula{color:#c678dd;}
                .hljs-section,.hljs-name,.hljs-selector-tag,.hljs-deletion,.hljs-subst{color:#e06c75;}
                .hljs-literal{color:#56b6c2;}
                .hljs-string,.hljs-regexp,.hljs-addition,.hljs-attribute,.hljs-meta-string{color:#98c379;}
                .hljs-built_in,.hljs-class .hljs-title{color:#e6c07b;}
                .hljs-attr,.hljs-variable,.hljs-template-variable,.hljs-type,.hljs-selector-class,
                .hljs-selector-attr,.hljs-selector-pseudo,.hljs-number{color:#d19a66;}
                .hljs-symbol,.hljs-bullet,.hljs-link,.hljs-meta,.hljs-selector-id,.hljs-title{color:#61aeee;}
            }
        </style>
        \(mermaidInline)
        \(highlightInline)
        </head>
        <body>
        <button class="toggle-btn" onclick="toggle()">&lt;/&gt;</button>
        <div id="preview"></div>
        <pre id="source"></pre>
        <script>
        var rawPreview = `\(jsEscaped)`;
        var rawSource = `\(jsEscapedSource)`;
        var showing = 'preview';
        function toggle() {
            var p = document.getElementById('preview');
            var s = document.getElementById('source');
            var btn = document.querySelector('.toggle-btn');
            if (showing === 'preview') {
                p.style.display = 'none';
                s.style.display = 'block';
                btn.innerHTML = '\u{1f441}';
                showing = 'source';
            } else {
                s.style.display = 'none';
                p.style.display = 'block';
                btn.innerHTML = '&lt;/&gt;';
                showing = 'preview';
            }
        }
        function md(s) {
            var mermaidBlocks = [];
            s = s.replace(/```(\\w*)\\n([\\s\\S]*?)```/g, function(_, lang, code) {
                if (lang === 'mermaid') {
                    var idx = mermaidBlocks.length;
                    mermaidBlocks.push(code.trim());
                    return '<div data-mermaid-placeholder="' + idx + '"></div>';
                }
                var cls = lang ? ' class="language-' + lang + '"' : '';
                return '<pre><code' + cls + '>' + esc(code.trim()) + '</code></pre>';
            });
            s = s.replace(/^\\|(.+)\\|\\n\\|[-| :]+\\|\\n((?:\\|.+\\|\\n?)*)/gm, function(_, header, body) {
                var ths = header.split('|').map(function(c){return '<th>'+c.trim()+'</th>';}).join('');
                var rows = body.trim().split('\\n').map(function(r){
                    return '<tr>'+r.replace(/^\\||\\|$/g,'').split('|').map(function(c){return '<td>'+c.trim()+'</td>';}).join('')+'</tr>';
                }).join('');
                return '<table><thead><tr>'+ths+'</tr></thead><tbody>'+rows+'</tbody></table>';
            });
            s = s.replace(/^######\\s+(.*)$/gm, '<h6>$1</h6>');
            s = s.replace(/^#####\\s+(.*)$/gm, '<h5>$1</h5>');
            s = s.replace(/^####\\s+(.*)$/gm, '<h4>$1</h4>');
            s = s.replace(/^###\\s+(.*)$/gm, '<h3>$1</h3>');
            s = s.replace(/^##\\s+(.*)$/gm, '<h2>$1</h2>');
            s = s.replace(/^#\\s+(.*)$/gm, '<h1>$1</h1>');
            s = s.replace(/^---+$/gm, '<hr>');
            s = s.replace(/^>\\s+(.*)$/gm, '<blockquote>$1</blockquote>');
            s = s.replace(/\\*\\*\\*(.+?)\\*\\*\\*/g, '<b><i>$1</i></b>');
            s = s.replace(/\\*\\*(.+?)\\*\\*/g, '<b>$1</b>');
            s = s.replace(/\\*(.+?)\\*/g, '<i>$1</i>');
            s = s.replace(/`([^`]+)`/g, '<code>$1</code>');
            s = s.replace(/!\\[([^\\]]*)\\]\\(([^)]+)\\)/g, '<img alt="$1" src="$2">');
            s = s.replace(/\\[([^\\]]+)\\]\\(([^)]+)\\)/g, '<a href="$2">$1</a>');
            s = s.replace(/^[\\-\\*]\\s+(.*)$/gm, '<li>$1</li>');
            s = s.replace(/((?:<li>.*<\\/li>\\n?)+)/g, '<ul>$1</ul>');
            s = s.replace(/^\\d+\\.\\s+(.*)$/gm, '<li>$1</li>');
            s = s.replace(/^(?!<[hupoltbd]|<li|<bl|<hr|<im|<a )(.+)$/gm, '<p>$1</p>');
            s = s.replace(/<\\/blockquote>\\n<blockquote>/g, '<br>');
            s = s.replace(/<div data-mermaid-placeholder="(\\d+)"><\\/div>/g, function(_, idx) {
                return '<div class="mermaid">' + esc(mermaidBlocks[+idx]) + '</div>';
            });
            return s;
        }
        function esc(s) {
            return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        }
        document.getElementById('preview').innerHTML = md(rawPreview);
        document.getElementById('source').textContent = rawSource;
        if (window.hljs) { hljs.highlightAll(); }
        if (window.mermaid) {
            var isDark = matchMedia('(prefers-color-scheme: dark)').matches;
            mermaid.initialize({ startOnLoad: false, theme: isDark ? 'dark' : 'default' });
            mermaid.run({ querySelector: '.mermaid' });
        }
        </script>
        </body>
        </html>
        """

        let baseURL = URL(fileURLWithPath: filePath).deletingLastPathComponent()
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: baseURL)
        }
    }

    // MARK: - Video File

    private static func ffmpegPath() -> String? {
        if let bundled = Bundle.main.path(forResource: "ffmpeg", ofType: nil) {
            return bundled
        }
        let candidates = [
            "/opt/homebrew/bin/ffmpeg",
            "/usr/local/bin/ffmpeg",
            (NSHomeDirectory() as NSString).appendingPathComponent(".local/bin/ffmpeg"),
        ]
        return candidates.first { FileManager.default.isExecutableFile(atPath: $0) }
    }

    private func loadVideoFile(_ filePath: String) {
        let ext = URL(fileURLWithPath: filePath).pathExtension.lowercased()
        if Self.nativeVideoExtensions.contains(ext) {
            playVideoNative(filePath)
        } else {
            transcodeAndPlay(filePath)
        }
    }

    private func transcodeAndPlay(_ filePath: String) {
        let videoURL = URL(fileURLWithPath: (filePath as NSString).standardizingPath)
        let filename = videoURL.lastPathComponent
        let filenameEsc = filename
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")

        guard let ffmpeg = Self.ffmpegPath() else {
            let html = """
            <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
            body{background:#000;color:#aaa;font-family:-apple-system,sans-serif;
                 display:flex;flex-direction:column;justify-content:center;align-items:center;
                 height:100vh;gap:12px;margin:0;}
            code{color:#f87171;font-size:13px;}
            </style></head><body>
            <div>无法播放 <b>\(filenameEsc)</b></div>
            <div>需要 ffmpeg：<code>brew install ffmpeg</code></div>
            </body></html>
            """
            DispatchQueue.main.async { [weak self] in
                self?.webView.loadHTMLString(html, baseURL: nil)
            }
            return
        }

        let transcodingHtml = """
        <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
        body{background:#000;color:#aaa;font-family:-apple-system,sans-serif;
             display:flex;flex-direction:column;justify-content:center;align-items:center;
             height:100vh;gap:12px;margin:0;}
        .spinner{width:32px;height:32px;border:3px solid #333;border-top-color:#888;
                 border-radius:50%;animation:spin 0.8s linear infinite;}
        @keyframes spin{to{transform:rotate(360deg);}}
        </style></head><body>
        <div class="spinner"></div>
        <div>正在转码 <b>\(filenameEsc)</b>…</div>
        </body></html>
        """
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(transcodingHtml, baseURL: nil)
        }

        let outPath = (NSTemporaryDirectory() as NSString)
            .appendingPathComponent("anyview-transcode-\(videoURL.deletingPathExtension().lastPathComponent).mp4")

        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            guard let self else { return }
            let task = Process()
            task.executableURL = URL(fileURLWithPath: ffmpeg)
            task.arguments = ["-y", "-i", filePath,
                              "-c:v", "copy", "-c:a", "aac", "-movflags", "faststart",
                              outPath]
            let errPipe = Pipe()
            task.standardError = errPipe
            task.standardOutput = Pipe()

            self.transcodeProcess = task
            do { try task.run() } catch {
                DispatchQueue.main.async { self.showError("ffmpeg 启动失败: \(error)") }
                return
            }
            task.waitUntilExit()
            self.transcodeProcess = nil

            if FileManager.default.fileExists(atPath: outPath) {
                DispatchQueue.main.async { self.playVideoNative(outPath) }
            } else {
                let errData = errPipe.fileHandleForReading.readDataToEndOfFile()
                let errMsg = String(data: errData, encoding: .utf8) ?? "未知错误"
                let html = """
                <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
                body{background:#000;color:#aaa;font-family:-apple-system,sans-serif;
                     padding:24px;margin:0;}
                pre{color:#f87171;font-size:11px;white-space:pre-wrap;margin-top:12px;}
                </style></head><body>
                <div>转码失败：\(filenameEsc)</div>
                <pre>\(errMsg.replacingOccurrences(of: "<", with: "&lt;"))</pre>
                </body></html>
                """
                DispatchQueue.main.async { [weak self] in
                    self?.webView.loadHTMLString(html, baseURL: nil)
                }
            }
        }
    }

    private func playVideoNative(_ filePath: String) {
        let videoURL = URL(fileURLWithPath: (filePath as NSString).standardizingPath)
        let videoFileURI = videoURL.absoluteString
        let filename = videoURL.lastPathComponent
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")

        let html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { background: #000; display: flex; flex-direction: column;
                   height: 100vh; overflow: hidden; }
            video { flex: 1; width: 100%; min-height: 0; background: #000; display: block; }
            #toolbar { flex: 0 0 auto; display: flex; align-items: center; gap: 8px;
                       padding: 6px 12px; background: #111;
                       font-family: -apple-system, sans-serif; font-size: 12px; color: #aaa; }
            #toolbar button {
                padding: 4px 10px; border: 1px solid rgba(255,255,255,0.15);
                border-radius: 5px; background: rgba(255,255,255,0.08); color: #ccc;
                font-size: 12px; cursor: pointer; font-family: inherit; }
            #toolbar button:hover { background: rgba(255,255,255,0.15); }
            #sub-label { color: #666; font-style: italic; }
            #err { color: #f87171; font-size: 12px; display: none; padding: 4px 8px; }
        </style>
        </head>
        <body>
        <video id="v" controls preload="metadata">
            <source src="\(videoFileURI)">
            <track id="sub-track" kind="subtitles" label="字幕" srclang="und" default>
        </video>
        <div id="toolbar">
            <button onclick="document.getElementById('f').click()">加载字幕…</button>
            <span id="sub-label">未加载字幕</span>
            <span id="err"></span>
            <input type="file" id="f" accept=".srt,.vtt,.ass,.ssa,.sub,.sbv" style="display:none">
        </div>
        <script>
        function srtToVtt(s) {
            return 'WEBVTT\\n\\n' + s
                .replace(/\\r\\n/g, '\\n')
                .replace(/\\r/g, '\\n')
                .replace(/(\\d{2}:\\d{2}:\\d{2}),(\\d{3})/g, '$1.$2')
                .trim();
        }
        function assToVtt(s) {
            var vtt = 'WEBVTT\\n\\n';
            var idx = 1;
            s.split('\\n').forEach(function(line) {
                var m = line.match(/^Dialogue:\\s*\\d+,([\\d:.]+),([\\d:.]+),[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,(.*)/);
                if (!m) return;
                function t(ts) {
                    var p = ts.split(':');
                    return (p[0].length < 2 ? '0' + p[0] : p[0]) + ':' + p[1] + ':' + p[2].replace('.', '.');
                }
                var text = m[3].replace(/\\{[^}]*\\}/g, '').replace(/<[^>]+>/g, '');
                vtt += (idx++) + '\\n' + t(m[1]) + ' --> ' + t(m[2]) + '\\n' + text + '\\n\\n';
            });
            return vtt;
        }
        document.getElementById('f').onchange = function(e) {
            var file = e.target.files[0];
            if (!file) return;
            var err = document.getElementById('err');
            err.style.display = 'none';
            var reader = new FileReader();
            reader.onerror = function() { err.textContent = '读取失败'; err.style.display = 'inline'; };
            reader.onload = function(ev) {
                var content = ev.target.result;
                var ext = file.name.split('.').pop().toLowerCase();
                var vtt;
                try {
                    if (ext === 'vtt') { vtt = content; }
                    else if (ext === 'ass' || ext === 'ssa') { vtt = assToVtt(content); }
                    else { vtt = srtToVtt(content); }
                } catch(ex) {
                    err.textContent = '解析失败: ' + ex.message;
                    err.style.display = 'inline'; return;
                }
                var blob = new Blob([vtt], { type: 'text/vtt' });
                var url = URL.createObjectURL(blob);
                var track = document.getElementById('sub-track');
                var old = track.src;
                track.src = url;
                if (old && old.startsWith('blob:')) URL.revokeObjectURL(old);
                // Force track reload
                var v = document.getElementById('v');
                for (var i = 0; i < v.textTracks.length; i++) {
                    v.textTracks[i].mode = 'showing';
                }
                document.getElementById('sub-label').textContent = file.name;
            };
            reader.readAsText(file, 'UTF-8');
        };
        document.getElementById('v').onerror = function() {
            document.getElementById('err').textContent =
                '无法播放 \(filename)（编解码器不支持或文件损坏）';
            document.getElementById('err').style.display = 'inline';
        };
        </script>
        </body>
        </html>
        """

        let tmpDir = (NSTemporaryDirectory() as NSString).appendingPathComponent("anyview-video-tmp")
        try? FileManager.default.createDirectory(atPath: tmpDir, withIntermediateDirectories: true)
        let htmlPath = (tmpDir as NSString).appendingPathComponent("index.html")
        try? html.write(toFile: htmlPath, atomically: true, encoding: .utf8)

        let accessDir = videoURL.deletingLastPathComponent()
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadFileURL(URL(fileURLWithPath: htmlPath), allowingReadAccessTo: accessDir)
        }
    }

    // MARK: - Subtitle File

    private func loadSubtitleFile(_ filePath: String) {
        guard let raw = try? String(contentsOfFile: filePath, encoding: .utf8)
                ?? String(contentsOfFile: filePath, encoding: .isoLatin1) else {
            showError("Failed to read file")
            return
        }

        let sourceEscaped = raw
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
        let jsSource = raw
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "`", with: "\\`")
            .replacingOccurrences(of: "$", with: "\\$")
        let ext = URL(fileURLWithPath: filePath).pathExtension.lowercased()

        let html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <meta name="color-scheme" content="light dark">
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: -apple-system, sans-serif; font-size: 14px;
                   background: #fff; color: #1a1a1a; }
            #preview { padding: 0 0 40px; }
            .entry { display: flex; gap: 0; border-bottom: 1px solid #f0f0f0; }
            .entry:hover { background: #f9fafb; }
            .num { flex: 0 0 48px; padding: 10px 8px 10px 16px;
                   color: #aaa; font-size: 12px; font-variant-numeric: tabular-nums;
                   text-align: right; user-select: none; }
            .time { flex: 0 0 200px; padding: 10px 12px;
                    font-family: "SF Mono", Menlo, monospace; font-size: 11px;
                    color: #888; white-space: nowrap; }
            .text { flex: 1; padding: 10px 16px 10px 0; line-height: 1.5; }
            .text em { font-style: italic; }
            .text b { font-weight: bold; }
            .header { position: sticky; top: 0; z-index: 10;
                      padding: 8px 16px; font-size: 12px; color: #888;
                      background: #fff; border-bottom: 1px solid #e5e7eb;
                      display: flex; gap: 12px; }
            .header span { flex: 0 0 48px; text-align: right; }
            .header .th-time { flex: 0 0 200px; padding-left: 12px; }
            .header .th-text { flex: 1; }
            #source { display: none; margin: 0; padding: 20px 24px;
                      white-space: pre-wrap; word-wrap: break-word;
                      font-family: "SF Mono", Menlo, monospace; font-size: 13px;
                      line-height: 1.5; tab-size: 4; min-height: 100vh; }
            .toggle-btn { position: fixed; top: 12px; right: 16px; z-index: 9999;
                          width: 32px; height: 32px; border: none; border-radius: 50%;
                          background: rgba(0,0,0,0.06); color: #555;
                          font-size: 14px; cursor: pointer;
                          backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
                          display: flex; align-items: center; justify-content: center;
                          transition: all 0.2s; }
            .toggle-btn:hover { background: rgba(0,0,0,0.12); transform: scale(1.08); }
            @media (prefers-color-scheme: dark) {
                body { background: #1a1a1a; color: #d4d4d4; }
                .entry { border-bottom-color: #2a2a2a; }
                .entry:hover { background: #222; }
                .header { background: #1a1a1a; border-bottom-color: #333; }
                .toggle-btn { background: rgba(255,255,255,0.1); color: #aaa; }
                .toggle-btn:hover { background: rgba(255,255,255,0.18); }
            }
        </style>
        </head>
        <body>
        <button class="toggle-btn" onclick="toggle()">&lt;/&gt;</button>
        <div id="preview">
          <div class="header">
            <span>#</span><span class="th-time">Timecode</span><span class="th-text">Text</span>
          </div>
          <div id="entries"></div>
        </div>
        <pre id="source">\(sourceEscaped)</pre>
        <script>
        var showing = 'preview';
        function toggle() {
            var p = document.getElementById('preview');
            var s = document.getElementById('source');
            var btn = document.querySelector('.toggle-btn');
            if (showing === 'preview') {
                p.style.display = 'none'; s.style.display = 'block';
                btn.innerHTML = '\\u{1f441}'; showing = 'source';
            } else {
                s.style.display = 'none'; p.style.display = 'block';
                btn.innerHTML = '&lt;/&gt;'; showing = 'preview';
            }
        }
        function esc(s) {
            return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        }
        function inlineStyle(s) {
            return s.replace(/<i>(.*?)<\\/i>/g,'<em>$1</em>')
                    .replace(/<b>(.*?)<\\/b>/g,'<b>$1</b>')
                    .replace(/<[^>]+>/g,'');
        }
        var raw = `\(jsSource)`;
        var ext = '\(ext)';
        var entries = [];
        if (ext === 'vtt') {
            var blocks = raw.replace(/\\r\\n/g,'\\n').split(/\\n\\s*\\n/);
            blocks.forEach(function(b) {
                b = b.trim();
                if (!b || b === 'WEBVTT') return;
                var lines = b.split('\\n');
                var ti = lines.findIndex(function(l){ return l.indexOf('-->') !== -1; });
                if (ti === -1) return;
                var num = ti > 0 ? lines[0] : String(entries.length + 1);
                var time = lines[ti];
                var text = lines.slice(ti + 1).join('<br>');
                entries.push({ num: num, time: time, text: text });
            });
        } else {
            // SRT and fallback: blocks split on blank lines
            var blocks = raw.replace(/\\r\\n/g,'\\n').split(/\\n\\s*\\n/);
            blocks.forEach(function(b) {
                b = b.trim();
                if (!b) return;
                var lines = b.split('\\n');
                if (lines.length < 2) return;
                var num = lines[0].trim();
                var time = lines[1].trim();
                if (time.indexOf('-->') === -1) return;
                var text = lines.slice(2).join('<br>');
                entries.push({ num: num, time: time, text: text });
            });
        }
        var container = document.getElementById('entries');
        if (entries.length === 0) {
            container.innerHTML = '<div style="padding:24px 16px;color:#888;font-size:13px;">No subtitle entries found — see source view.</div>';
        } else {
            container.innerHTML = entries.map(function(e) {
                return '<div class="entry"><div class="num">' + esc(e.num) + '</div>' +
                       '<div class="time">' + esc(e.time) + '</div>' +
                       '<div class="text">' + inlineStyle(e.text) + '</div></div>';
            }).join('');
        }
        </script>
        </body>
        </html>
        """

        let baseURL = URL(fileURLWithPath: filePath).deletingLastPathComponent()
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: baseURL)
        }
    }

    // MARK: - TeX File

    private static func tectonicPath() -> String? {
        if let bundled = Bundle.main.path(forResource: "tectonic", ofType: nil) {
            return bundled
        }
        let candidates = [
            "/opt/homebrew/bin/tectonic",
            "/usr/local/bin/tectonic",
            (NSHomeDirectory() as NSString).appendingPathComponent(".local/bin/tectonic"),
        ]
        return candidates.first { FileManager.default.isExecutableFile(atPath: $0) }
    }

    private func loadTexFile(_ filePath: String) {
        guard let source = try? String(contentsOfFile: filePath, encoding: .utf8) else {
            showError("Failed to read file")
            return
        }
        let escaped = source
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
        let highlightInline = Self.highlightScript.isEmpty ? "" : "<script>\(Self.highlightScript)</script>"
        let latexGrammar = Self.hljsLatexScript.isEmpty ? "" : "<script>\(Self.hljsLatexScript)</script>"

        func makeSourceHtml(statusMsg: String, isError: Bool) -> String {
            let statusColor = isError ? "#b91c1c" : "#888"
            let statusHtml = statusMsg.isEmpty ? "" : """
            <div style="position:fixed;top:12px;right:16px;z-index:9999;padding:6px 12px;
                        font:12px -apple-system,sans-serif;color:\(statusColor);
                        background:rgba(0,0,0,0.06);border-radius:4px;max-width:60%;
                        white-space:pre-wrap;">\(statusMsg.replacingOccurrences(of: "\n", with: "<br>"))</div>
            """
            return """
            <!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="color-scheme" content="light dark">
            <style>
            body{margin:0;padding:20px 24px;font-family:"SF Mono",Menlo,monospace;font-size:13px;
                 background:#f8f9fa;color:#1a1a1a;}
            pre{margin:0;line-height:1.5;white-space:pre-wrap;word-wrap:break-word;tab-size:4;}
            .hljs{display:block;overflow-x:auto;padding:0;color:#333;background:transparent;}
            .hljs-comment,.hljs-quote{color:#998;font-style:italic;}
            .hljs-keyword,.hljs-selector-tag,.hljs-subst{color:#333;font-weight:bold;}
            .hljs-number,.hljs-literal,.hljs-variable,.hljs-template-variable,.hljs-tag .hljs-attr{color:#008080;}
            .hljs-string,.hljs-doctag{color:#d14;}
            .hljs-title,.hljs-section,.hljs-selector-id{color:#900;font-weight:bold;}
            .hljs-built_in,.hljs-builtin-name{color:#0086b3;}
            .hljs-meta{color:#999;font-weight:bold;}
            @media(prefers-color-scheme:dark){
                body{background:#1e1e1e;color:#d4d4d4;}
                .hljs{color:#abb2bf;}
                .hljs-comment,.hljs-quote{color:#5c6370;font-style:italic;}
                .hljs-keyword,.hljs-formula{color:#c678dd;}
                .hljs-string,.hljs-regexp,.hljs-addition,.hljs-attribute{color:#98c379;}
                .hljs-built_in{color:#e6c07b;}
                .hljs-number,.hljs-type,.hljs-selector-class,.hljs-selector-pseudo{color:#d19a66;}
            }
            </style>
            \(highlightInline)
            \(latexGrammar)
            </head><body>
            \(statusHtml)
            <pre><code class="language-latex">\(escaped)</code></pre>
            <script>if(window.hljs){hljs.highlightAll();}</script>
            </body></html>
            """
        }

        // Subfile check — skip compilation if no \begin{document}
        guard source.contains("\\begin{document}") else {
            let html = makeSourceHtml(statusMsg: "Subfile (no \\begin{document}) — source only", isError: false)
            DispatchQueue.main.async { [weak self] in
                self?.webView.loadHTMLString(html, baseURL: URL(fileURLWithPath: filePath))
            }
            return
        }

        // Show source + "Compiling…" immediately
        let compilingHtml = makeSourceHtml(statusMsg: "Compiling…", isError: false)
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(compilingHtml, baseURL: URL(fileURLWithPath: filePath))
        }

        guard let tectonic = Self.tectonicPath() else {
            let html = makeSourceHtml(statusMsg: "tectonic not found — install with: brew install tectonic", isError: true)
            DispatchQueue.main.async { [weak self] in
                self?.webView.loadHTMLString(html, baseURL: URL(fileURLWithPath: filePath))
            }
            return
        }

        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            guard let self else { return }

            let tmpDir = (NSTemporaryDirectory() as NSString)
                .appendingPathComponent("anyview-tex-\(UUID().uuidString)")
            try? FileManager.default.createDirectory(atPath: tmpDir, withIntermediateDirectories: true)

            let task = Process()
            task.executableURL = URL(fileURLWithPath: tectonic)
            task.arguments = ["-Z", "continue-on-errors", "--outdir", tmpDir, filePath]
            let pipe = Pipe()
            task.standardOutput = pipe
            task.standardError = pipe

            do {
                try task.run()
                task.waitUntilExit()
            } catch {
                let html = makeSourceHtml(statusMsg: "Failed to launch tectonic: \(error.localizedDescription)", isError: true)
                DispatchQueue.main.async { self.webView.loadHTMLString(html, baseURL: URL(fileURLWithPath: filePath)) }
                try? FileManager.default.removeItem(atPath: tmpDir)
                return
            }

            let baseName = URL(fileURLWithPath: filePath).deletingPathExtension().lastPathComponent
            let pdfPath = (tmpDir as NSString).appendingPathComponent("\(baseName).pdf")

            if FileManager.default.fileExists(atPath: pdfPath) {
                self.tempDir = tmpDir
                self.texPdfPath = pdfPath
                self.texSourceHtml = makeSourceHtml(statusMsg: "", isError: false)
                DispatchQueue.main.async { self.showTexPdf() }
            } else {
                let errData = pipe.fileHandleForReading.readDataToEndOfFile()
                let errMsg = String(data: errData, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines) ?? "Compilation failed"
                let html = makeSourceHtml(statusMsg: errMsg, isError: true)
                DispatchQueue.main.async { self.webView.loadHTMLString(html, baseURL: URL(fileURLWithPath: filePath)) }
                try? FileManager.default.removeItem(atPath: tmpDir)
            }
        }
    }

    // MARK: - Code File

    private static let langMap: [String: String] = [
        "swift": "swift", "cs": "csharp", "py": "python", "js": "javascript",
        "ts": "typescript", "tsx": "typescript", "jsx": "javascript",
        "go": "go", "rs": "rust", "rb": "ruby", "java": "java",
        "kt": "kotlin", "scala": "scala", "c": "c", "h": "c",
        "cpp": "cpp", "hpp": "cpp", "m": "objectivec", "mm": "objectivec",
        "lua": "lua", "r": "r", "pl": "perl", "php": "php", "dart": "dart",
        "zig": "zig", "nim": "nim", "ex": "elixir", "exs": "elixir",
        "erl": "erlang", "hs": "haskell", "ml": "ocaml", "fs": "fsharp",
        "asm": "x86asm", "s": "x86asm", "sql": "sql",
        "sh": "bash", "bash": "bash", "zsh": "bash", "fish": "shell",
        "bat": "dos", "ps1": "powershell", "cmd": "dos",
        "xml": "xml", "json": "json", "yaml": "yaml", "yml": "yaml",
        "toml": "ini", "ini": "ini", "cfg": "ini", "conf": "ini",
        "plist": "xml", "graphql": "graphql", "proto": "protobuf",
        "css": "css", "scss": "scss", "sass": "scss", "less": "less",
        "svg": "xml", "rst": "plaintext", "txt": "plaintext",
        "log": "plaintext", "diff": "diff", "patch": "diff",
        "makefile": "makefile", "cmake": "cmake",
        "html": "html", "htm": "html",
        "tex": "latex", "sty": "latex", "cls": "latex", "bib": "tex", "bbl": "tex",
    ]

    private func loadCodeFile(_ filePath: String) {
        guard let content = try? String(contentsOfFile: filePath, encoding: .utf8) else {
            showError("Failed to read file")
            return
        }

        let escaped = content
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")

        let lineCount = content.components(separatedBy: "\n").count
        let ext = URL(fileURLWithPath: filePath).pathExtension.lowercased()
        let lang = Self.langMap[ext] ?? ""
        let langClass = lang.isEmpty ? "" : "language-\(lang)"
        let escapedFilename = URL(fileURLWithPath: filePath).lastPathComponent
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")

        let highlightInline = Self.highlightScript.isEmpty ? "" : "<script>\(Self.highlightScript)</script>"

        let html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <meta name="color-scheme" content="light dark">
        <style>
            body { margin: 0; padding: 20px 24px;
                   font-family: "SF Mono", Menlo, Consolas, monospace; font-size: 13px;
                   background: #f8f9fa; color: #1a1a1a; }
            pre { margin: 0; line-height: 1.5; white-space: pre-wrap; word-wrap: break-word;
                  tab-size: 4; }
            pre code { font-family: inherit; font-size: inherit; }
            .header { color: #888; font-size: 11px; margin-bottom: 12px; padding-bottom: 8px;
                      border-bottom: 1px solid #ddd; }
            .hljs{display:block;overflow-x:auto;padding:0;color:#333;background:transparent;}
            .hljs-comment,.hljs-quote{color:#998;font-style:italic;}
            .hljs-keyword,.hljs-selector-tag,.hljs-subst{color:#333;font-weight:bold;}
            .hljs-number,.hljs-literal,.hljs-variable,.hljs-template-variable,.hljs-tag .hljs-attr{color:#008080;}
            .hljs-string,.hljs-doctag{color:#d14;}
            .hljs-title,.hljs-section,.hljs-selector-id{color:#900;font-weight:bold;}
            .hljs-subst{font-weight:normal;}
            .hljs-type,.hljs-class .hljs-title{color:#458;font-weight:bold;}
            .hljs-tag,.hljs-name,.hljs-attribute{color:#000080;font-weight:normal;}
            .hljs-regexp,.hljs-link{color:#009926;}
            .hljs-symbol,.hljs-bullet{color:#990073;}
            .hljs-built_in,.hljs-builtin-name{color:#0086b3;}
            .hljs-meta{color:#999;font-weight:bold;}
            .hljs-deletion{background:#fdd;}.hljs-addition{background:#dfd;}
            @media (prefers-color-scheme: dark) {
                body { background: #1e1e1e; color: #d4d4d4; }
                .header { border-bottom-color: #333; }
                .hljs{color:#abb2bf;background:transparent;}
                .hljs-comment,.hljs-quote{color:#5c6370;font-style:italic;}
                .hljs-doctag,.hljs-keyword,.hljs-formula{color:#c678dd;}
                .hljs-section,.hljs-name,.hljs-selector-tag,.hljs-deletion,.hljs-subst{color:#e06c75;}
                .hljs-literal{color:#56b6c2;}
                .hljs-string,.hljs-regexp,.hljs-addition,.hljs-attribute,.hljs-meta-string{color:#98c379;}
                .hljs-built_in,.hljs-class .hljs-title{color:#e6c07b;}
                .hljs-attr,.hljs-variable,.hljs-template-variable,.hljs-type,.hljs-selector-class,
                .hljs-selector-attr,.hljs-selector-pseudo,.hljs-number{color:#d19a66;}
                .hljs-symbol,.hljs-bullet,.hljs-link,.hljs-meta,.hljs-selector-id,.hljs-title{color:#61aeee;}
            }
        </style>
        \(highlightInline)
        </head>
        <body>
        <div class="header">\(lineCount) lines · \(escapedFilename)</div>
        <pre><code class="\(langClass)">\(escaped)</code></pre>
        <script>if(window.hljs){hljs.highlightAll();}</script>
        </body>
        </html>
        """

        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: URL(fileURLWithPath: filePath))
        }
    }

    // MARK: - Docx / Docmod / Doct

    private func loadDocxContent(_ filePath: String) {
        guard let data = try? Data(contentsOf: URL(fileURLWithPath: filePath)) else {
            showError("Failed to read docx file")
            return
        }
        let base64 = data.base64EncodedString()

        let html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="UTF-8">
        <meta name="color-scheme" content="light">
        <style>
            * { box-sizing: border-box; }
            html, body { margin: 0; padding: 0; background: #e5e7eb; }
            #container { padding: 24px 0; min-height: 100vh; }
            .docx-wrapper { background: transparent !important; padding: 0 !important; }
            .docx-wrapper > section.docx { box-shadow: 0 2px 8px rgba(0,0,0,0.12); margin: 0 auto 16px; }
            #status { position: fixed; top: 12px; right: 16px; z-index: 9999;
                      padding: 6px 12px; font: 12px -apple-system, sans-serif;
                      background: rgba(0,0,0,0.7); color: #fff; border-radius: 4px;
                      backdrop-filter: blur(8px); }
            #status:empty { display: none; }
            #status.error { background: #b91c1c; }
            /* Map Windows CJK fonts to macOS equivalents (normal + bold) */
            @font-face { font-family: '宋体'; font-weight: normal; src: local('STSong-Light'), local('STSong'); }
            @font-face { font-family: '宋体'; font-weight: bold; src: local('STSong'); }
            @font-face { font-family: 'SimSun'; font-weight: normal; src: local('STSong-Light'), local('STSong'); }
            @font-face { font-family: 'SimSun'; font-weight: bold; src: local('STSong'); }
            @font-face { font-family: '微软雅黑'; font-weight: normal; src: local('PingFang SC'), local('PingFang SC Regular'); }
            @font-face { font-family: '微软雅黑'; font-weight: bold; src: local('PingFang SC Semibold'), local('PingFang SC Medium'); }
            @font-face { font-family: 'Microsoft YaHei'; font-weight: normal; src: local('PingFang SC'), local('PingFang SC Regular'); }
            @font-face { font-family: 'Microsoft YaHei'; font-weight: bold; src: local('PingFang SC Semibold'), local('PingFang SC Medium'); }
            @font-face { font-family: '黑体'; font-weight: normal; src: local('STHeiti'); }
            @font-face { font-family: '黑体'; font-weight: bold; src: local('STHeiti Medium'); }
            @font-face { font-family: 'SimHei'; font-weight: normal; src: local('STHeiti'); }
            @font-face { font-family: 'SimHei'; font-weight: bold; src: local('STHeiti Medium'); }
            @font-face { font-family: '楷体'; src: local('STKaiti'); }
            @font-face { font-family: 'KaiTi'; src: local('STKaiti'); }
            @font-face { font-family: '仿宋'; src: local('STFangsong'); }
            @font-face { font-family: 'FangSong'; src: local('STFangsong'); }
            @font-face { font-family: '等线'; font-weight: normal; src: local('PingFang SC'); }
            @font-face { font-family: '等线'; font-weight: bold; src: local('PingFang SC Semibold'); }
            @font-face { font-family: 'DengXian'; font-weight: normal; src: local('PingFang SC'); }
            @font-face { font-family: 'DengXian'; font-weight: bold; src: local('PingFang SC Semibold'); }
        </style>
        <script>\(Self.jszipScript)</script>
        <script>\(Self.docxPreviewScript)</script>
        </head>
        <body>
        <div id="status">加载中…</div>
        <div id="container"></div>
        <script>
        (async function() {
            var status = document.getElementById('status');
            var container = document.getElementById('container');
            var b64 = "\(base64)";
            var bin = atob(b64);
            var bytes = new Uint8Array(bin.length);
            for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            var MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            var opts = {
                className: 'docx', inWrapper: true, breakPages: true,
                ignoreLastRenderedPageBreak: true, experimental: true,
                trimXmlDeclaration: true, renderHeaders: true, renderFooters: true,
                renderFootnotes: true, renderEndnotes: true,
                renderComments: true, renderChanges: true,
            };
            try {
                await docx.renderAsync(new Blob([bytes], { type: MIME }), container, null, opts);
                status.textContent = '';
            } catch (e1) {
                // Retry after stripping OMML math (common pandoc-generated docx issue)
                try {
                    status.textContent = '正在处理数学公式…';
                    var zip = await JSZip.loadAsync(bytes);
                    var docFile = zip.file('word/document.xml');
                    if (!docFile) throw e1;
                    var xml = await docFile.async('string');
                    var cleaned = xml
                        .replace(/<m:oMathPara\\b[^>]*>[\\s\\S]*?<\\/m:oMathPara>/g,
                                 '<w:p><w:r><w:t>[公式]</w:t></w:r></w:p>')
                        .replace(/<m:oMath\\b[^>]*>[\\s\\S]*?<\\/m:oMath>/g,
                                 '<w:r><w:t>[公式]</w:t></w:r>');
                    zip.file('word/document.xml', cleaned);
                    var newBytes = await zip.generateAsync({ type: 'uint8array' });
                    container.innerHTML = '';
                    await docx.renderAsync(new Blob([newBytes], { type: MIME }), container, null, opts);
                    status.textContent = '数学公式已转为占位符';
                    setTimeout(function() { status.textContent = ''; }, 3000);
                } catch (e2) {
                    status.className = 'error';
                    status.textContent = '渲染失败: ' + (e2 && e2.message ? e2.message : e2);
                }
            }
        })();
        </script>
        </body>
        </html>
        """

        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: nil)
        }
    }

    private func loadDocmodContent(_ filePath: String, extractedDir: String) {
        let html: String
        do {
            html = try DocmodCLI.render(filePath: filePath)
        } catch let error as DocmodCLI.CLIError {
            showError(error.errorDescription ?? "docmod not found")
            return
        } catch {
            showError("Unexpected error: \(error.localizedDescription)")
            return
        }

        let baseURL = URL(fileURLWithPath: extractedDir)
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: baseURL)
        }
    }

    private func loadDoctContent(_ filePath: String, extractedDir: String) {
        let fm = FileManager.default

        let metaJson = readJsonFile(path: extractedDir + "/meta.json")
        let formatJson = readJsonFile(path: extractedDir + "/format.json")
        let contentJson = readJsonFile(path: extractedDir + "/content.json")

        let guidesDir = extractedDir + "/guides"
        let guides = (try? fm.contentsOfDirectory(atPath: guidesDir)) ?? []

        var documentPreview = ""
        let sourceDocx = extractedDir + "/source.docx"
        let renderTarget = fm.fileExists(atPath: sourceDocx) ? sourceDocx : filePath
        do {
            documentPreview = try DocmodCLI.render(filePath: renderTarget)
            if let bodyStart = documentPreview.range(of: "<body>"),
               let bodyEnd = documentPreview.range(of: "</body>") {
                documentPreview = String(documentPreview[bodyStart.upperBound..<bodyEnd.lowerBound])
            }
        } catch {
            documentPreview = "<p style='color:#999'>Preview not available</p>"
        }

        let name = jsonString(metaJson, key: "name") ?? URL(fileURLWithPath: filePath).lastPathComponent
        let description = jsonString(metaJson, key: "description") ?? ""

        let html = buildDoctHtml(
            name: name,
            description: description,
            metaJson: metaJson,
            formatJson: formatJson,
            contentJson: contentJson,
            guides: guides,
            documentPreview: documentPreview
        )

        let baseURL = URL(fileURLWithPath: extractedDir)
        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(html, baseURL: baseURL)
        }
    }

    private func buildDoctHtml(
        name: String,
        description: String,
        metaJson: String?,
        formatJson: String?,
        contentJson: String?,
        guides: [String],
        documentPreview: String
    ) -> String {
        let esc = { (s: String) -> String in
            s.replacingOccurrences(of: "&", with: "&amp;")
             .replacingOccurrences(of: "<", with: "&lt;")
             .replacingOccurrences(of: ">", with: "&gt;")
        }

        var sections = ""

        if let fj = formatJson, !fj.isEmpty {
            sections += """
            <details open>
            <summary>format.json</summary>
            <pre><code>\(esc(prettyJson(fj)))</code></pre>
            </details>
            """
        }

        if let cj = contentJson, !cj.isEmpty, cj != "{}" {
            sections += """
            <details open>
            <summary>content.json</summary>
            <pre><code>\(esc(prettyJson(cj)))</code></pre>
            </details>
            """
        }

        if !guides.isEmpty {
            let guideList = guides.map { "  <li>\(esc($0))</li>" }.joined(separator: "\n")
            sections += """
            <details>
            <summary>Guides (\(guides.count))</summary>
            <ul>
            \(guideList)
            </ul>
            </details>
            """
        }

        if let mj = metaJson, !mj.isEmpty {
            sections += """
            <details>
            <summary>meta.json</summary>
            <pre><code>\(esc(prettyJson(mj)))</code></pre>
            </details>
            """
        }

        return """
        <!DOCTYPE html>
        <html lang="zh">
        <head>
        <meta charset="UTF-8">
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: -apple-system, "Helvetica Neue", sans-serif; color: #1a1a1a; background: #f5f5f5; }
            .header { padding: 32px 40px 24px; background: #1a1a1a; color: white; }
            .header h1 { font-size: 24px; font-weight: 600; margin-bottom: 4px; }
            .header .badge { display: inline-block; font-size: 11px; font-weight: 500; padding: 2px 8px;
                             background: rgba(255,255,255,0.15); border-radius: 4px; margin-right: 8px;
                             vertical-align: middle; letter-spacing: 0.5px; }
            .header p { font-size: 14px; color: rgba(255,255,255,0.7); margin-top: 8px; }
            .content { max-width: 800px; margin: 0 auto; padding: 24px 40px 60px; }
            details { background: white; border-radius: 8px; margin-bottom: 12px;
                      box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
            summary { padding: 14px 20px; font-size: 14px; font-weight: 600; cursor: pointer;
                      user-select: none; }
            summary:hover { color: #2563eb; }
            details pre { padding: 0 20px 16px; overflow-x: auto; }
            details code { font-family: "SF Mono", Menlo, monospace; font-size: 12px; line-height: 1.6;
                           color: #334155; white-space: pre-wrap; word-break: break-word; }
            details ul { padding: 0 20px 16px 40px; font-size: 13px; line-height: 1.8; }
            .preview-section { background: white; border-radius: 8px; margin-bottom: 12px;
                               box-shadow: 0 1px 3px rgba(0,0,0,0.08); overflow: hidden; }
            .preview-label { padding: 14px 20px; font-size: 14px; font-weight: 600; }
            .preview-frame { border-top: 1px solid #e5e7eb; padding: 24px; }
        </style>
        </head>
        <body>
        <div class="header">
            <h1><span class="badge">.doct</span> \(esc(name))</h1>
            \(description.isEmpty ? "" : "<p>\(esc(description))</p>")
        </div>
        <div class="content">
            <div class="preview-section">
                <div class="preview-label">Document Preview</div>
                <div class="preview-frame">\(documentPreview)</div>
            </div>
            \(sections)
        </div>
        </body>
        </html>
        """
    }

    // MARK: - JSON Helpers

    private func readJsonFile(path: String) -> String? {
        try? String(contentsOfFile: path, encoding: .utf8)
    }

    private func jsonString(_ json: String?, key: String) -> String? {
        guard let json, let data = json.data(using: .utf8),
              let obj = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
              let value = obj[key] as? String else { return nil }
        return value
    }

    private func prettyJson(_ raw: String) -> String {
        guard let data = raw.data(using: .utf8),
              let obj = try? JSONSerialization.jsonObject(with: data),
              let pretty = try? JSONSerialization.data(withJSONObject: obj, options: [.prettyPrinted, .sortedKeys]),
              let str = String(data: pretty, encoding: .utf8) else { return raw }
        return str
    }

    // MARK: - Error Display

    private func showError(_ message: String) {
        let escapedMessage = message
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\n", with: "<br>")

        let errorHtml = """
        <!DOCTYPE html>
        <html>
        <head><meta charset="utf-8"></head>
        <body style="font-family: -apple-system, sans-serif; padding: 40px; color: #333;">
        <h2 style="color: #c00;">Error</h2>
        <p style="white-space: pre-wrap;">\(escapedMessage)</p>
        </body>
        </html>
        """

        DispatchQueue.main.async { [weak self] in
            self?.webView.loadHTMLString(errorHtml, baseURL: nil)
        }
    }
}
