export function resolveZipPath(basePath: string, target: string): string {
  if (!target) return "";

  if (!basePath.includes("/")) return target.replace(/^\/+/, "");

  const baseParts = basePath.split("/");
  baseParts.pop();

  const targetParts = target.split("/");
  for (const part of targetParts) {
    if (!part || part === ".") continue;
    if (part === "..") {
      if (baseParts.length) baseParts.pop();
    } else {
      baseParts.push(part);
    }
  }

  return baseParts.join("/");
}

export function getImageMimeType(path: string): string {
  const lower = path.toLowerCase();

  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".webp")) return "image/webp";
  if (lower.endsWith(".bmp")) return "image/bmp";
  if (lower.endsWith(".svg")) return "image/svg+xml";

  return "application/octet-stream";
}

export function toBlobUrl(
  data: Uint8Array | ArrayBuffer,
  mimeType: string
): string {
  let arrayBuffer: ArrayBuffer;

  if (data instanceof Uint8Array) {
    arrayBuffer = data.buffer.slice(
      data.byteOffset,
      data.byteOffset + data.byteLength
    ) as ArrayBuffer;
  } else {
    arrayBuffer = data;
  }

  const blob = new Blob([arrayBuffer], { type: mimeType });
  return URL.createObjectURL(blob);
}