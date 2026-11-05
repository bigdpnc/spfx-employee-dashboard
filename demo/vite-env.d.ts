/// <reference types="vite/client" />

// SPFx generates typed `*.module.scss.ts` files at build time; this covers the demo build.
declare module '*.module.scss' {
  const classes: { readonly [key: string]: string };
  export default classes;
}
