import "react";

declare module "react" {
  interface InputHTMLAttributes<T> {
    webkitdirectory?: string;
    directory?: string;
  }
}

declare namespace JSX {
  interface IntrinsicElements {
    input: any;
  }
}
