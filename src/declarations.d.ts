// src/declarations.d.ts
declare module "*.png";
declare module "*.jpg";
declare module "*.jpeg";
declare module "*.svg";

declare module "*.webm" {
  const src: string;
  export default src;
}

declare module "*.mp4" {
  const src: string;
  export default src;
}