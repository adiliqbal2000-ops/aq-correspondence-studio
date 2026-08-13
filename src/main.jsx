import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// Catches errors that happen during React's render (as opposed to errors
// during initial script/module loading, which index.html's inline script
// handles separately). Shows the error directly on screen so it can be
// read and reported from a phone, without needing a desktop console.
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null, info: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    this.setState({ info });
  }

  render() {
    if (this.state.error) {
      return (
        <div
          style={{
            fontFamily: "monospace",
            whiteSpace: "pre-wrap",
            wordBreak: "break-word",
            padding: 16,
            margin: 16,
            color: "#7f1d1d",
            background: "#fef2f2",
            border: "1px solid #fca5a5",
            borderRadius: 8,
            fontSize: 13,
            lineHeight: 1.5,
          }}
        >
          <strong>The app crashed while rendering:</strong>
          {"\n\n"}
          {this.state.error && (this.state.error.message || String(this.state.error))}
          {this.state.info && this.state.info.componentStack ? "\n\nComponent stack:" + this.state.info.componentStack : ""}
        </div>
      );
    }
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>
);
