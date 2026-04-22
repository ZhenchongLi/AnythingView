use adw::prelude::*;
use std::cell::RefCell;
use std::path::{Path, PathBuf};
use std::rc::Rc;

use crate::drop_target::attach_drop_target;
use crate::file_watcher::FileWatcher;
use crate::renderer::{Renderer, RendererFactory};

pub struct ViewerWindow {
    window: adw::ApplicationWindow,
    file_path: PathBuf,
    renderer: RefCell<Option<Box<dyn Renderer>>>,
    zoom_level: RefCell<f64>,
    zoom_label: gtk::Label,
    on_close: RefCell<Option<Box<dyn Fn()>>>,
    on_open_files: RefCell<Option<Box<dyn Fn(Vec<PathBuf>)>>>,
    _watcher: RefCell<Option<FileWatcher>>,
}

impl ViewerWindow {
    pub fn new(app: &adw::Application, path: &Path) -> Rc<Self> {
        let filename = path
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("AnyView")
            .to_string();

        let window = adw::ApplicationWindow::builder()
            .application(app)
            .title(&filename)
            .default_width(900)
            .default_height(1000)
            .build();

        let ext = path
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_string();
        let renderer = RendererFactory::renderer_for(&ext);

        let header = adw::HeaderBar::new();
        let zoom_out = gtk::Button::from_icon_name("zoom-out-symbolic");
        zoom_out.set_tooltip_text(Some("Zoom Out (Ctrl+-)"));
        let zoom_label = gtk::Label::new(Some("100%"));
        zoom_label.set_width_chars(5);
        let zoom_reset = gtk::Button::builder().child(&zoom_label).build();
        zoom_reset.set_tooltip_text(Some("Reset Zoom (Ctrl+0)"));
        let zoom_in = gtk::Button::from_icon_name("zoom-in-symbolic");
        zoom_in.set_tooltip_text(Some("Zoom In (Ctrl+=)"));

        let zoom_box = gtk::Box::new(gtk::Orientation::Horizontal, 2);
        zoom_box.add_css_class("linked");
        zoom_box.append(&zoom_out);
        zoom_box.append(&zoom_reset);
        zoom_box.append(&zoom_in);
        header.pack_start(&zoom_box);

        let content_box = gtk::Box::new(gtk::Orientation::Vertical, 0);
        content_box.append(&header);

        let renderer_widget = renderer.widget();
        renderer_widget.set_hexpand(true);
        renderer_widget.set_vexpand(true);
        content_box.append(&renderer_widget);

        window.set_content(Some(&content_box));

        let this = Rc::new(Self {
            window,
            file_path: path.to_path_buf(),
            renderer: RefCell::new(Some(renderer)),
            zoom_level: RefCell::new(1.0),
            zoom_label,
            on_close: RefCell::new(None),
            on_open_files: RefCell::new(None),
            _watcher: RefCell::new(None),
        });

        // Zoom button handlers
        {
            let this_weak = Rc::downgrade(&this);
            zoom_in.connect_clicked(move |_| {
                if let Some(s) = this_weak.upgrade() {
                    s.zoom_by(0.1);
                }
            });
        }
        {
            let this_weak = Rc::downgrade(&this);
            zoom_out.connect_clicked(move |_| {
                if let Some(s) = this_weak.upgrade() {
                    s.zoom_by(-0.1);
                }
            });
        }
        {
            let this_weak = Rc::downgrade(&this);
            zoom_reset.connect_clicked(move |_| {
                if let Some(s) = this_weak.upgrade() {
                    s.set_zoom(1.0);
                }
            });
        }

        // Keyboard shortcuts: Ctrl++, Ctrl+-, Ctrl+0, Ctrl+R
        let controller = gtk::EventControllerKey::new();
        {
            let this_weak = Rc::downgrade(&this);
            controller.connect_key_pressed(move |_, key, _, state| {
                let Some(s) = this_weak.upgrade() else {
                    return glib::Propagation::Proceed;
                };
                let ctrl = state.contains(gdk::ModifierType::CONTROL_MASK);
                if !ctrl {
                    return glib::Propagation::Proceed;
                }
                match key {
                    gdk::Key::plus | gdk::Key::equal | gdk::Key::KP_Add => {
                        s.zoom_by(0.1);
                        glib::Propagation::Stop
                    }
                    gdk::Key::minus | gdk::Key::KP_Subtract => {
                        s.zoom_by(-0.1);
                        glib::Propagation::Stop
                    }
                    gdk::Key::_0 | gdk::Key::KP_0 => {
                        s.set_zoom(1.0);
                        glib::Propagation::Stop
                    }
                    gdk::Key::r | gdk::Key::R => {
                        s.reload();
                        glib::Propagation::Stop
                    }
                    _ => glib::Propagation::Proceed,
                }
            });
        }
        this.window.add_controller(controller);

        // Drag-and-drop
        {
            let this_weak = Rc::downgrade(&this);
            attach_drop_target(this.window.upcast_ref::<gtk::Widget>(), move |paths| {
                if let Some(s) = this_weak.upgrade() {
                    if let Some(cb) = s.on_open_files.borrow().as_ref() {
                        cb(paths);
                    }
                }
            });
        }

        // Close handler
        {
            let this_weak = Rc::downgrade(&this);
            this.window.connect_close_request(move |_| {
                if let Some(s) = this_weak.upgrade() {
                    *s._watcher.borrow_mut() = None;
                    if let Some(cb) = s.on_close.borrow().as_ref() {
                        cb();
                    }
                }
                glib::Propagation::Proceed
            });
        }

        // Initial load + file watcher
        this.reload();
        this.start_watching();

        this
    }

    pub fn file_path(&self) -> &Path {
        &self.file_path
    }

    pub fn present(&self) {
        self.window.present();
    }

    pub fn set_on_close<F: Fn() + 'static>(&self, cb: F) {
        *self.on_close.borrow_mut() = Some(Box::new(cb));
    }

    pub fn set_on_open_files<F: Fn(Vec<PathBuf>) + 'static>(&self, cb: F) {
        *self.on_open_files.borrow_mut() = Some(Box::new(cb));
    }

    fn reload(&self) {
        if let Some(r) = self.renderer.borrow().as_ref() {
            r.load(&self.file_path);
        }
    }

    fn start_watching(self: &Rc<Self>) {
        let weak = Rc::downgrade(self);
        let watcher = FileWatcher::new(&self.file_path, move || {
            if let Some(s) = weak.upgrade() {
                s.reload();
            }
        });
        *self._watcher.borrow_mut() = watcher;
    }

    fn zoom_by(&self, delta: f64) {
        let new_level = *self.zoom_level.borrow() + delta;
        self.set_zoom(new_level);
    }

    fn set_zoom(&self, level: f64) {
        let snapped = (level * 10.0).round() / 10.0;
        let clamped = snapped.clamp(0.5, 3.0);
        *self.zoom_level.borrow_mut() = clamped;
        if let Some(r) = self.renderer.borrow().as_ref() {
            r.set_zoom(clamped);
        }
        self.zoom_label
            .set_text(&format!("{}%", (clamped * 100.0).round() as i32));
    }
}
