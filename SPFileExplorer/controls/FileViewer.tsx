import * as React from "react";
import { Modal } from "@fluentui/react/lib/Modal";
import { IconButton } from "@fluentui/react/lib/Button";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { IFileSystemItem } from "./IFileSystemItem";
import { IStrings, formatString } from "./Strings";
import { FileItemIcon } from "./FileItemIcon";
import {
  ViewerKind,
  getBrowserUrl,
  getDownloadUrl,
  getEmbedPreviewUrl,
  getPdfFrameUrl,
  getPreviewImageUrl,
  getViewerImageUrl,
  getViewerKind,
  isSafeViewerUrl,
  toEncodedUrl,
} from "./FileLinks";
import { IZoomPanSurfaceHandle, ZoomPanSurface } from "./ZoomPanSurface";
import { ViewerIcon, ViewerIconName } from "./ViewerIcons";

/** Frames (PDF, Office, text) don't always report that they've loaded: stop the spinner after this. */
const FRAME_SPINNER_MS = 6000;

export interface IFileViewerProps {
  /** Files that can be stepped through with Previous/Next (the current folder or search results). */
  items: IFileSystemItem[];
  /** Key of the file shown. */
  currentKey: string;
  /** Returns the SharePoint URL of a file, or undefined if it's not known. */
  getFileUrl: (item: IFileSystemItem) => string | undefined;
  /** URL of the SharePoint site, used for previews of files the browser can't show itself. */
  siteUrl?: string;
  strings: IStrings;
  onNavigate: (key: string) => void;
  onClose: () => void;
  /** Opens a URL in a new tab (or downloads), the same way as the explorer's file actions. */
  openUrl: (url: string) => void;
}

interface ISource {
  kind: ViewerKind;
  /** What is actually loaded: "image" also covers the first-page image of a PDF in fallback mode. */
  render: "image" | "frame" | "video" | "audio" | "none";
  url: string;
}

const sourceFor = (
  item: IFileSystemItem,
  fileUrl: string | undefined,
  siteUrl: string | undefined,
  pdfAsImage: boolean
): ISource => {
  const kind = getViewerKind(item.filetype);
  const none: ISource = { kind, render: "none", url: "" };
  if (!fileUrl || !isSafeViewerUrl(fileUrl)) return none;
  switch (kind) {
    case "image":
      return siteUrl || /^(jpe?g|png|gif|bmp|webp|svg|ico|avif)$/i.test(item.filetype ?? "")
        ? { kind, render: "image", url: getViewerImageUrl(siteUrl ?? "", fileUrl, item.filetype) }
        : none;
    case "pdf":
      return pdfAsImage && siteUrl
        ? { kind, render: "image", url: getPreviewImageUrl(siteUrl, fileUrl, 6) }
        : { kind, render: "frame", url: getPdfFrameUrl(fileUrl) };
    case "office": {
      const url = siteUrl ? getEmbedPreviewUrl(siteUrl, fileUrl, item.filetype) : undefined;
      return url ? { kind, render: "frame", url } : none;
    }
    case "video":
      return { kind, render: "video", url: toEncodedUrl(fileUrl) };
    case "audio":
      return { kind, render: "audio", url: toEncodedUrl(fileUrl) };
    case "text":
      return { kind, render: "frame", url: toEncodedUrl(fileUrl) };
    default:
      return none;
  }
};

/**
 * Full-screen viewer for SharePoint files: images (zoom, pan, rotate), PDFs (browser PDF viewer, with zoom and
 * hand tool around it, or the page as an image), Office documents (Office for the web), video, audio and text.
 * Everything is loaded by the browser from SharePoint with the user's own SharePoint session.
 */
export const FileViewer = (props: IFileViewerProps) => {
  const { items, currentKey, strings } = props;
  const index = Math.max(0, items.findIndex((i) => i.key === currentKey));
  const item = items[index];
  const fileUrl = item ? props.getFileUrl(item) : undefined;
  const [pdfAsImage, setPdfAsImage] = React.useState(false);
  const [handTool, setHandTool] = React.useState(false);
  const [scale, setScale] = React.useState(1);
  const [natural, setNatural] = React.useState<{ width: number; height: number }>();
  const [loading, setLoading] = React.useState(true);
  const [failed, setFailed] = React.useState(false);
  const surface = React.useRef<IZoomPanSurfaceHandle>(null);

  const source = item ? sourceFor(item, fileUrl, props.siteUrl, pdfAsImage) : undefined;

  // Reset per file.
  React.useEffect(() => {
    setNatural(undefined);
    setLoading(true);
    setFailed(false);
    setHandTool(false);
  }, [currentKey, pdfAsImage]);
  React.useEffect(() => setPdfAsImage(false), [currentKey]);
  React.useEffect(() => {
    if (source?.render !== "frame") return;
    const t = window.setTimeout(() => setLoading(false), FRAME_SPINNER_MS);
    return () => window.clearTimeout(t);
  }, [currentKey, source?.render, source?.url]);

  const go = (delta: number) => {
    if (items.length < 2) return;
    const next = items[(index + delta + items.length) % items.length];
    props.onNavigate(next.key);
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    // Left/Right step through files unless the zoom surface (which pans with arrows) has focus.
    const inSurface = (e.target as HTMLElement).closest?.(".zoomPanSurface");
    if (!inSurface && e.key === "ArrowRight") { go(1); e.preventDefault(); }
    if (!inSurface && e.key === "ArrowLeft") { go(-1); e.preventDefault(); }
    if (e.key === "PageDown") { go(1); e.preventDefault(); }
    if (e.key === "PageUp") { go(-1); e.preventDefault(); }
  };

  if (!item || !source) return null;

  const zoomable = source.render === "image" || source.render === "frame";
  const title = item.fullname;
  const position = items.length > 1 ? formatString(strings.ViewerPosition, String(index + 1), String(items.length)) : "";

  const content = (() => {
    switch (source.render) {
      case "image":
        return (
          <ZoomPanSurface
            ref={surface}
            mode="image"
            naturalSize={natural}
            handTool
            resetKey={`${currentKey}|${source.url}`}
            ariaLabel={formatString(strings.ViewerImageLabel, title)}
            onScaleChange={setScale}
          >
            <img
              className="viewerImage"
              src={source.url}
              alt={title}
              draggable={false}
              referrerPolicy="no-referrer"
              onLoad={(e) => {
                const img = e.currentTarget;
                setNatural({ width: img.naturalWidth, height: img.naturalHeight });
                setLoading(false);
              }}
              onError={() => { setLoading(false); setFailed(true); }}
            />
          </ZoomPanSurface>
        );
      case "frame":
        return (
          <ZoomPanSurface
            ref={surface}
            mode="frame"
            handTool={handTool}
            resetKey={`${currentKey}|${source.url}`}
            ariaLabel={formatString(strings.ViewerDocumentLabel, title)}
            onScaleChange={setScale}
          >
            <iframe
              key={source.url}
              className="viewerFrame"
              src={source.url}
              title={title}
              referrerPolicy="no-referrer"
              sandbox="allow-scripts allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-forms allow-downloads"
              onLoad={() => setLoading(false)}
            />
          </ZoomPanSurface>
        );
      case "video":
        return (
          <div className="viewerMedia">
            <video key={source.url} src={source.url} controls autoPlay={false} onLoadedData={() => setLoading(false)}
              onError={() => { setLoading(false); setFailed(true); }} />
          </div>
        );
      case "audio":
        return (
          <div className="viewerMedia audio">
            <FileItemIcon isFolder={false} extension={item.filetype} size="large" className="tileIcon large" strings={strings} />
            <audio key={source.url} src={source.url} controls onLoadedData={() => setLoading(false)}
              onError={() => { setLoading(false); setFailed(true); }} />
          </div>
        );
      default:
        return (
          <div className="viewerUnsupported">
            <FileItemIcon isFolder={false} extension={item.filetype} size="large" className="tileIcon large" strings={strings} />
            <p>{fileUrl ? strings.ViewerNotSupported : strings.ViewerNoLink}</p>
            {fileUrl && (
              <div className="viewerUnsupportedActions">
                <DefaultButton iconProps={{ iconName: "Globe" }} text={strings.OpenInBrowser} onClick={() => props.openUrl(getBrowserUrl(fileUrl))} />
                <DefaultButton iconProps={{ iconName: "Download" }} text={strings.Download} onClick={() => props.openUrl(getDownloadUrl(fileUrl))} />
              </div>
            )}
          </div>
        );
    }
  })();

  const tb = (key: string, icon: ViewerIconName, label: string, onClick: () => void, opts: { disabled?: boolean; checked?: boolean } = {}) => (
    <IconButton
      key={key}
      className="viewerToolbarButton"
      onRenderIcon={() => <ViewerIcon name={icon} />}
      title={label}
      ariaLabel={label}
      disabled={opts.disabled}
      checked={opts.checked}
      toggle={opts.checked !== undefined}
      onClick={onClick}
    />
  );

  return (
    <Modal
      isOpen
      onDismiss={props.onClose}
      isBlocking={false}
      className="fileViewerModal"
      containerClassName="fileViewerContainer"
      titleAriaId="fileViewerTitle"
    >
      <div className="fileViewer" onKeyDown={onKeyDown}>
        <div className="viewerToolbar" role="toolbar" aria-label={strings.ViewerToolbar}>
          <div className="viewerTitle">
            <span id="fileViewerTitle" title={title}>{title}</span>
            {position && <span className="viewerPosition">{position}</span>}
          </div>
          <div className="viewerTools">
            {items.length > 1 && tb("prev", "ChevronLeft", strings.ViewerPrevious, () => go(-1))}
            {items.length > 1 && tb("next", "ChevronRight", strings.ViewerNext, () => go(1))}
            {zoomable && (
              <>
                <span className="viewerDivider" />
                {tb("zoomOut", "ZoomOut", strings.ZoomOut, () => surface.current?.zoomOut())}
                <span className="viewerZoom" aria-live="polite">{Math.round(scale * 100)}%</span>
                {tb("zoomIn", "ZoomIn", strings.ZoomIn, () => surface.current?.zoomIn())}
                {tb("fit", "ArrowMaximize", strings.ZoomFit, () => surface.current?.fit())}
                {source.render === "image" && tb("actual", "ScaleFill", strings.ZoomActualSize, () => surface.current?.actualSize())}
                {source.render === "image" && tb("rotate", "ArrowRotateClockwise", strings.Rotate, () => surface.current?.rotate())}
                {source.render === "frame" && tb("hand", "HandLeft", strings.HandTool, () => setHandTool((h) => !h), { checked: handTool })}
              </>
            )}
            {source.kind === "pdf" && props.siteUrl && (
              tb("pdfMode", pdfAsImage ? "DocumentPdf" : "Image", pdfAsImage ? strings.PdfShowDocument : strings.PdfShowAsImage,
                () => setPdfAsImage((v) => !v), { checked: pdfAsImage })
            )}
            <span className="viewerDivider" />
            {fileUrl && tb("browser", "Globe", strings.OpenInBrowser, () => props.openUrl(getBrowserUrl(fileUrl)))}
            {fileUrl && tb("download", "ArrowDownload", strings.Download, () => props.openUrl(getDownloadUrl(fileUrl)))}
            {tb("close", "Dismiss", strings.Close, props.onClose)}
          </div>
        </div>
        <div className="viewerBody">
          {content}
          {loading && source.render !== "none" && !failed && (
            <div className="viewerLoading"><Spinner size={SpinnerSize.large} label={strings.Loading} /></div>
          )}
          {failed && (
            <div className="viewerFailed" role="alert">
              <p>{strings.ViewerLoadFailed}</p>
              {fileUrl && <DefaultButton iconProps={{ iconName: "Globe" }} text={strings.OpenInBrowser} onClick={() => props.openUrl(getBrowserUrl(fileUrl))} />}
            </div>
          )}
        </div>
        {source.render === "frame" && (
          <div className="viewerHint">{source.kind === "pdf" ? strings.ViewerPdfHint : strings.ViewerFrameHint}</div>
        )}
        {source.render === "image" && source.kind === "pdf" && (
          <div className="viewerHint">{strings.ViewerPdfImageHint}</div>
        )}
      </div>
    </Modal>
  );
};
