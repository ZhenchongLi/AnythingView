import Cocoa

protocol FindBarViewDelegate: AnyObject {
    func findBar(_ bar: FindBarView, didSearch query: String, forward: Bool)
    func findBarDidRequestClose(_ bar: FindBarView)
}

final class FindBarView: NSView, NSTextFieldDelegate {
    static let preferredHeight: CGFloat = 32

    weak var delegate: FindBarViewDelegate?

    private let textField = NSTextField()
    private let statusLabel = NSTextField(labelWithString: "")
    private let prevButton = NSButton()
    private let nextButton = NSButton()
    private let closeButton = NSButton()
    private let separator = NSBox()

    override init(frame: NSRect) {
        super.init(frame: frame)
        autoresizingMask = [.width]
        setupUI()
    }

    required init?(coder: NSCoder) { fatalError("init(coder:) not supported") }

    var query: String { textField.stringValue }

    func focusInput() {
        window?.makeFirstResponder(textField)
        textField.currentEditor()?.selectAll(nil)
    }

    func setStatus(_ text: String, isError: Bool = false) {
        statusLabel.stringValue = text
        statusLabel.textColor = isError ? .systemRed : .secondaryLabelColor
    }

    private func setupUI() {
        wantsLayer = true
        layer?.backgroundColor = NSColor.windowBackgroundColor.cgColor

        textField.placeholderString = "Find"
        textField.target = self
        textField.action = #selector(onSubmitNext)
        textField.delegate = self
        textField.font = .systemFont(ofSize: 13)
        textField.translatesAutoresizingMaskIntoConstraints = false

        statusLabel.font = .systemFont(ofSize: 11)
        statusLabel.textColor = .secondaryLabelColor
        statusLabel.translatesAutoresizingMaskIntoConstraints = false
        statusLabel.setContentHuggingPriority(.defaultHigh, for: .horizontal)

        configureIconButton(prevButton, symbol: "chevron.up", tooltip: "Previous (⇧↵)", action: #selector(onPrev))
        configureIconButton(nextButton, symbol: "chevron.down", tooltip: "Next (↵)", action: #selector(onNext))
        configureIconButton(closeButton, symbol: "xmark", tooltip: "Close (⎋)", action: #selector(onClose))

        let stack = NSStackView(views: [textField, statusLabel, prevButton, nextButton, closeButton])
        stack.orientation = .horizontal
        stack.spacing = 8
        stack.edgeInsets = NSEdgeInsets(top: 4, left: 12, bottom: 4, right: 12)
        stack.alignment = .centerY
        stack.translatesAutoresizingMaskIntoConstraints = false
        addSubview(stack)

        separator.boxType = .separator
        separator.translatesAutoresizingMaskIntoConstraints = false
        addSubview(separator)

        NSLayoutConstraint.activate([
            stack.leadingAnchor.constraint(equalTo: leadingAnchor),
            stack.trailingAnchor.constraint(equalTo: trailingAnchor),
            stack.topAnchor.constraint(equalTo: topAnchor),
            stack.bottomAnchor.constraint(equalTo: separator.topAnchor),
            separator.leadingAnchor.constraint(equalTo: leadingAnchor),
            separator.trailingAnchor.constraint(equalTo: trailingAnchor),
            separator.bottomAnchor.constraint(equalTo: bottomAnchor),
            separator.heightAnchor.constraint(equalToConstant: 1),
            textField.widthAnchor.constraint(equalToConstant: 240),
        ])
    }

    private func configureIconButton(_ button: NSButton, symbol: String, tooltip: String, action: Selector) {
        button.image = NSImage(systemSymbolName: symbol, accessibilityDescription: tooltip)
        button.bezelStyle = .texturedRounded
        button.isBordered = false
        button.imagePosition = .imageOnly
        button.toolTip = tooltip
        button.target = self
        button.action = action
        button.translatesAutoresizingMaskIntoConstraints = false
        button.widthAnchor.constraint(equalToConstant: 24).isActive = true
        button.heightAnchor.constraint(equalToConstant: 24).isActive = true
    }

    @objc private func onSubmitNext() { fireSearch(forward: true) }
    @objc private func onNext() { fireSearch(forward: true) }
    @objc private func onPrev() { fireSearch(forward: false) }
    @objc private func onClose() { delegate?.findBarDidRequestClose(self) }

    private func fireSearch(forward: Bool) {
        let q = textField.stringValue
        guard !q.isEmpty else {
            setStatus("")
            return
        }
        delegate?.findBar(self, didSearch: q, forward: forward)
    }

    func control(_ control: NSControl, textView: NSTextView, doCommandBy selector: Selector) -> Bool {
        if selector == #selector(NSResponder.cancelOperation(_:)) {
            delegate?.findBarDidRequestClose(self)
            return true
        }
        if selector == #selector(NSResponder.insertNewline(_:)) {
            let shiftHeld = NSApp.currentEvent?.modifierFlags.contains(.shift) ?? false
            fireSearch(forward: !shiftHeld)
            return true
        }
        return false
    }

    func controlTextDidChange(_ obj: Notification) {
        if textField.stringValue.isEmpty {
            setStatus("")
        }
    }
}
