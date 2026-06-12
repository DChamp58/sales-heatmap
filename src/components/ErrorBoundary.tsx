import { Component, type ErrorInfo, type ReactNode } from 'react';

interface Props {
  children: ReactNode;
}
interface State {
  error: Error | null;
  componentStack: string | null;
}

// Catches render/lifecycle errors anywhere in the tree and shows the actual
// message on screen instead of letting the app fail to a blank white page.
export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, componentStack: null };

  static getDerivedStateFromError(error: Error): Partial<State> {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    this.setState({ componentStack: info.componentStack ?? null });
    console.error('SalesHeatmap crashed:', error, info);
  }

  render() {
    const { error, componentStack } = this.state;
    if (!error) return this.props.children;

    return (
      <div style={{ maxWidth: '820px', margin: '40px auto', padding: '0 16px', fontFamily: 'system-ui' }}>
        <div style={{ background: '#FEF2F2', border: '1px solid #FECACA', borderRadius: '12px', padding: '20px' }}>
          <h2 style={{ margin: '0 0 8px', color: '#7F1D1D', fontSize: '18px' }}>The app hit an error</h2>
          <p style={{ margin: '0 0 12px', color: '#991B1B', fontSize: '14px' }}>
            Please copy the details below so the problem can be diagnosed, then reload the page.
          </p>
          <pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: '12px', color: '#991B1B', background: '#FFF', border: '1px solid #FECACA', borderRadius: '8px', padding: '12px', margin: 0 }}>
            {error.message}
            {error.stack ? `\n\n${error.stack}` : ''}
            {componentStack ? `\n\nComponent stack:${componentStack}` : ''}
          </pre>
          <button
            onClick={() => location.reload()}
            style={{ marginTop: '14px', background: '#F76902', border: 'none', color: '#fff', borderRadius: '8px', padding: '8px 16px', fontWeight: 600, cursor: 'pointer' }}
          >
            Reload
          </button>
        </div>
      </div>
    );
  }
}
