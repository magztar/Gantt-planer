import SwiftUI
import WebKit

struct BundledWebAppView: View {
    @State private var loadError: String?

    var body: some View {
        ZStack {
            HostedWebView(loadError: $loadError)

            if let loadError {
                ContentUnavailableView(
                    "Kunde inte ladda webbappen",
                    systemImage: "exclamationmark.triangle",
                    description: Text(loadError)
                )
                .padding(24)
                .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 16, style: .continuous))
                .padding(24)
            }
        }
    }
}

#if os(macOS)
struct HostedWebView: NSViewRepresentable {
    @Binding var loadError: String?

    func makeCoordinator() -> Coordinator {
        Coordinator(loadError: $loadError)
    }

    func makeNSView(context: Context) -> WKWebView {
        let configuration = WKWebViewConfiguration()
        configuration.defaultWebpagePreferences.allowsContentJavaScript = true

        let webView = WKWebView(frame: .zero, configuration: configuration)
        webView.navigationDelegate = context.coordinator
        webView.allowsBackForwardNavigationGestures = false
        webView.setValue(false, forKey: "drawsBackground")
        context.coordinator.loadBundledApp(in: webView)
        return webView
    }

    func updateNSView(_ webView: WKWebView, context: Context) {}
}
#else
struct HostedWebView: UIViewRepresentable {
    @Binding var loadError: String?

    func makeCoordinator() -> Coordinator {
        Coordinator(loadError: $loadError)
    }

    func makeUIView(context: Context) -> WKWebView {
        let configuration = WKWebViewConfiguration()
        configuration.defaultWebpagePreferences.allowsContentJavaScript = true

        let webView = WKWebView(frame: .zero, configuration: configuration)
        webView.navigationDelegate = context.coordinator
        webView.allowsBackForwardNavigationGestures = false
        context.coordinator.loadBundledApp(in: webView)
        return webView
    }

    func updateUIView(_ webView: WKWebView, context: Context) {}
}
#endif

final class Coordinator: NSObject, WKNavigationDelegate {
    private var loadError: Binding<String?>

    init(loadError: Binding<String?>) {
        self.loadError = loadError
    }

    func loadBundledApp(in webView: WKWebView) {
        guard let resourcesURL = Bundle.main.resourceURL else {
            loadError.wrappedValue = "Kunde inte läsa appens resurskatalog."
            return
        }

        let webAppDirectory = resourcesURL.appendingPathComponent("WebApp", isDirectory: true)
        guard FileManager.default.fileExists(atPath: webAppDirectory.path) else {
            loadError.wrappedValue = "Katalogen WebApp finns inte i app-bundlen. Kontrollera build-fasen som kopierar webbappen. Förväntad sökväg: \(webAppDirectory.path)"
            return
        }

        let indexURL = webAppDirectory.appendingPathComponent("index.html")
        guard FileManager.default.fileExists(atPath: indexURL.path) else {
            loadError.wrappedValue = "Filen index.html hittades inte i app-bundlen. Förväntad sökväg: \(indexURL.path)"
            return
        }

        loadError.wrappedValue = nil
        webView.loadFileURL(indexURL, allowingReadAccessTo: webAppDirectory)
    }

    func webView(_ webView: WKWebView, didFail navigation: WKNavigation!, withError error: Error) {
        loadError.wrappedValue = error.localizedDescription
    }

    func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        loadError.wrappedValue = error.localizedDescription
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        loadError.wrappedValue = nil
    }
}
