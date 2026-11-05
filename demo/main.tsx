import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DemoApp } from './DemoApp';
import './demo.css';

// React 17 render API, matching the React version SPFx ships with.
ReactDOM.render(
  <React.StrictMode>
    <DemoApp />
  </React.StrictMode>,
  document.getElementById('root'),
);
