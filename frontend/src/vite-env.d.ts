/// <reference types="vite/client" />

declare module "react-plotly.js" {
  import { Component } from "react";
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  export default class Plot extends Component<any> {}
}

// eslint-disable-next-line @typescript-eslint/no-namespace
declare namespace Plotly {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  type Data = any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  type Layout = any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  type Config = any;
}
