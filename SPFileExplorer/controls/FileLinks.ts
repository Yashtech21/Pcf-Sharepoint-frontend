/**
 * SharePoint URLs of the document location shown by the control.
 */
export interface ILocationUrls {
  /**
   * Absolute URL of the folder of the document location.
   */
  locationUrl: string;
  /**
   * URL of the SharePoint site (collection) of the document location.
   */
  siteUrl: string;
  /**
   * Path of the document location relative to its document library (path of the root folder).
   */
  rootPath: string;
}

const OFFICE_DESKTOP_SCHEMES: { [extension: string]: string } = {
  doc: "ms-word",
  docx: "ms-word",
  docm: "ms-word",
  dot: "ms-word",
  dotx: "ms-word",
  rtf: "ms-word",
  xls: "ms-excel",
  xlsx: "ms-excel",
  xlsm: "ms-excel",
  xlsb: "ms-excel",
  csv: "ms-excel",
  ppt: "ms-powerpoint",
  pptx: "ms-powerpoint",
  pptm: "ms-powerpoint",
  pps: "ms-powerpoint",
  ppsx: "ms-powerpoint",
  vsd: "ms-visio",
  vsdx: "ms-visio",
  mpp: "ms-project",
};

const PREVIEW_EXTENSIONS = [
  "jpg",
  "jpeg",
  "png",
  "gif",
  "bmp",
  "webp",
  "tif",
  "tiff",
  "heic",
  "pdf",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "ppt",
  "pptx",
  "rtf",
  "txt",
  "vsdx",
  "mp4",
];

// File types shown by Office for the web in the full preview.
const EMBED_PREVIEW_EXTENSIONS = [
  "doc",
  "docx",
  "docm",
  "dot",
  "dotx",
  "xls",
  "xlsx",
  "xlsm",
  "xlsb",
  "ppt",
  "pptx",
  "pptm",
  "pps",
  "ppsx",
  "vsd",
  "vsdx",
];

const encodePath = (path: string): string =>
  path
    .split("/")
    .filter((segment) => segment !== "")
    .map((segment) => encodeURIComponent(segment))
    .join("/");

const trimSlashes = (value: string): string => value.replace(/\/+$/, "");

/**
 * Returns the URL of the SharePoint site of an URL ("/sites/x" or "/teams/x", otherwise the root site).
 */
export const getSiteUrl = (url: string): string => {
  const match = /^(https?:\/\/[^/]+)(\/(?:sites|teams|personal)\/[^/]+)?/i.exec(url);
  return match ? `${match[1]}${match[2] ?? ""}` : url;
};

/**
 * Returns the path of a folder relative to the document location, e.g. "A/B" for "Contoso_1/A/B".
 */
export const getPathBelowRoot = (rootPath: string, folderPath: string): string => {
  if (!folderPath || folderPath === rootPath) {
    return "";
  }
  return folderPath.startsWith(`${rootPath}/`)
    ? folderPath.substring(rootPath.length + 1)
    : folderPath;
};

/**
 * Returns the absolute URL of a folder of the document location.
 */
export const getFolderUrl = (urls: ILocationUrls, folderPath: string): string => {
  const subPath = getPathBelowRoot(urls.rootPath, folderPath);
  return subPath
    ? `${trimSlashes(urls.locationUrl)}/${encodePath(subPath)}`
    : trimSlashes(urls.locationUrl);
};

/**
 * Returns the absolute URL of a file or folder.
 * @param absoluteUrl Absolute URL returned by Dataverse for the item, if any.
 * @param name File name.
 * @param folderPath Path of the folder that contains the item.
 */
export const getItemUrl = (
  urls: ILocationUrls,
  absoluteUrl: string | undefined,
  name: string,
  folderPath: string
): string =>
  absoluteUrl && /^https?:\/\//i.test(absoluteUrl)
    ? absoluteUrl
    : `${getFolderUrl(urls, folderPath)}/${encodeURIComponent(name)}`;

/**
 * Returns the path of the folder that contains an item, based on its relative location.
 * The relative location of a file can either be the path of its folder or its own path.
 * Dataverse returns the file name in the path URL-encoded for files in subfolders
 * (for example "Folder/Sub folder/My%20file.docx"), so both forms of the name are accepted.
 */
export const getContainingFolderPath = (
  relativeLocation: string,
  name: string
): string => {
  const location = trimSlashes(relativeLocation ?? "");
  for (const fileName of [name, encodeURIComponent(name)]) {
    if (location === fileName) {
      return "";
    }
    if (location.endsWith(`/${fileName}`)) {
      return location.substring(0, location.length - fileName.length - 1);
    }
  }
  return location;
};

/**
 * Returns the URL with the characters encoded that must be encoded in a link. Dataverse returns
 * URLs with unencoded spaces in folder names; characters that are already encoded are kept.
 */
export const toEncodedUrl = (url: string): string => {
  try {
    return new URL(url).href;
  } catch {
    return url;
  }
};

const withQuery = (url: string, query: string): string =>
  `${url}${url.indexOf("?") >= 0 ? "&" : "?"}${query}`;

/**
 * URL that opens a file in the browser (Office for the web for Office documents).
 */
export const getBrowserUrl = (fileUrl: string): string => withQuery(fileUrl, "web=1");

/**
 * URL that downloads a file.
 */
export const getDownloadUrl = (fileUrl: string): string =>
  withQuery(fileUrl, "download=1");

/**
 * Office URI that opens a file in the desktop application, or undefined for other file types.
 */
export const getDesktopAppUrl = (
  fileUrl: string,
  extension: string | undefined
): string | undefined => {
  const scheme = OFFICE_DESKTOP_SCHEMES[(extension ?? "").toLowerCase()];
  return scheme ? `${scheme}:ofe|u|${toEncodedUrl(fileUrl)}` : undefined;
};

/**
 * Indicates if SharePoint can render a preview image of a file with the specified extension.
 */
export const canPreview = (extension: string | undefined): boolean =>
  PREVIEW_EXTENSIONS.indexOf((extension ?? "").toLowerCase()) >= 0;

/**
 * URL of the preview image rendered by SharePoint (requires the user to be signed in to SharePoint).
 * @param resolution 0 (smallest) to 6 (largest).
 */
export const getPreviewImageUrl = (
  siteUrl: string,
  fileUrl: string,
  resolution: number
): string =>
  `${trimSlashes(siteUrl)}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(
    fileUrl
  )}&resolution=${resolution}`;

/**
 * URL of the Office for the web view of a file that can be shown in a frame (requires the user
 * to be signed in to SharePoint), or undefined for file types that Office for the web cannot show.
 */
export const getEmbedPreviewUrl = (
  siteUrl: string,
  fileUrl: string,
  extension: string | undefined
): string | undefined => {
  if (EMBED_PREVIEW_EXTENSIONS.indexOf((extension ?? "").toLowerCase()) < 0) {
    return undefined;
  }
  let path: string;
  try {
    path = decodeURIComponent(new URL(toEncodedUrl(fileUrl)).pathname);
  } catch {
    return undefined;
  }
  return `${trimSlashes(siteUrl)}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(
    path
  )}&action=embedview`;
};

// ── Built-in viewer ─────────────────────────────────────────────────────────────────────────────────
// Files are loaded by the browser straight from SharePoint with the user's own SharePoint session
// (no app registration). Media elements and frames don't need CORS, so no file content is downloaded
// by the control itself.

/** Kind of viewer used for a file in the full-screen viewer. */
export type ViewerKind = "image" | "pdf" | "office" | "video" | "audio" | "text" | "none";

// Formats every browser can draw natively in an <img>.
const NATIVE_IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg", "ico", "avif"];
// Image formats browsers can't draw: SharePoint converts them to a preview image.
const CONVERTED_IMAGE_EXTENSIONS = ["heic", "heif", "tif", "tiff"];
const VIDEO_EXTENSIONS = ["mp4", "m4v", "webm", "ogv", "mov"];
const AUDIO_EXTENSIONS = ["mp3", "wav", "m4a", "aac", "oga", "ogg", "flac", "opus"];
const TEXT_EXTENSIONS = [
  "txt", "log", "csv", "tsv", "json", "xml", "md", "yml", "yaml", "ini", "cfg", "conf",
  "sql", "js", "ts", "css", "cs", "java", "py", "ps1", "sh",
];

export const getViewerKind = (extension: string | undefined): ViewerKind => {
  const ext = (extension ?? "").toLowerCase();
  if (NATIVE_IMAGE_EXTENSIONS.indexOf(ext) >= 0 || CONVERTED_IMAGE_EXTENSIONS.indexOf(ext) >= 0) return "image";
  if (ext === "pdf") return "pdf";
  if (EMBED_PREVIEW_EXTENSIONS.indexOf(ext) >= 0) return "office";
  if (VIDEO_EXTENSIONS.indexOf(ext) >= 0) return "video";
  if (AUDIO_EXTENSIONS.indexOf(ext) >= 0) return "audio";
  if (TEXT_EXTENSIONS.indexOf(ext) >= 0) return "text";
  return "none";
};

/**
 * Source of the image shown in the viewer: the file itself for formats the browser can draw, otherwise
 * SharePoint's largest preview image (resolution 6).
 */
export const getViewerImageUrl = (siteUrl: string, fileUrl: string, extension: string | undefined): string =>
  NATIVE_IMAGE_EXTENSIONS.indexOf((extension ?? "").toLowerCase()) >= 0
    ? toEncodedUrl(fileUrl)
    : getPreviewImageUrl(siteUrl, fileUrl, 6);

/**
 * PDF shown with the browser's built-in PDF viewer in a frame (all pages, text selection, search).
 * The fragment asks the viewer to fit the page width.
 */
export const getPdfFrameUrl = (fileUrl: string): string => `${toEncodedUrl(fileUrl)}#view=FitH`;

/**
 * Only https links are loaded in the viewer (frames, images, media); anything else is refused.
 */
export const isSafeViewerUrl = (url: string | undefined): boolean => {
  if (!url) return false;
  try {
    return new URL(url).protocol === "https:";
  } catch {
    return false;
  }
};
