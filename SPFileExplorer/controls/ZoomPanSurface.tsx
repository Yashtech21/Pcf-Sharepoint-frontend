import * as React from "react";

export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 8;
const WHEEL_STEP = 1.15;
const KEY_PAN_PX = 60;

export const clampZoom = (z: number): number => Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z));

/** Transform state of the surface. `scale` is relative to the content's natural size (1 = 100%). */
export interface IZoomPanState {
  scale: number;
  x: number;
  y: number;
  rotation: number;
}

/**
 * Zoom around a point (relative to the surface centre) so that the content under that point stays put.
 */
export const zoomAt = (s: IZoomPanState, nextScale: number, px: number, py: number): IZoomPanState => {
  const scale = clampZoom(nextScale);
  const k = scale / s.scale;
  return { ...s, scale, x: px - (px - s.x) * k, y: py - (py - s.y) * k };
};

export interface IZoomPanSurfaceHandle {
  zoomIn(): void;
  zoomOut(): void;
  fit(): void;
  actualSize(): void;
  rotate(): void;
}

export interface IZoomPanSurfaceProps {
  /** "image": pan/zoom an image (fit to the surface). "frame": pan/zoom a frame that fills the surface. */
  mode: "image" | "frame";
  /** Natural size of the image, once loaded; used to fit it. */
  naturalSize?: { width: number; height: number };
  /**
   * Frames receive mouse and keyboard input themselves (scrolling PDF pages, Office). While the hand tool is on,
   * a transparent layer above the frame takes the input so it can be dragged.
   */
  handTool: boolean;
  /** Changes when another file is shown: the view is reset to "fit". */
  resetKey: string;
  ariaLabel: string;
  onScaleChange?: (scale: number) => void;
  children: React.ReactNode;
}

/**
 * Surface that shows its content with drag-to-pan and zoom: mouse wheel (zooms at the cursor), pinch on touch
 * screens, double-click (fit ↔ 200%), keyboard (+, -, 0, arrow keys) and the viewer's toolbar via the handle.
 */
export const ZoomPanSurface = React.forwardRef<IZoomPanSurfaceHandle, IZoomPanSurfaceProps>((props, ref) => {
  const { mode, naturalSize, handTool, resetKey, onScaleChange } = props;
  const surfaceRef = React.useRef<HTMLDivElement>(null);
  const [state, setState] = React.useState<IZoomPanState>({ scale: 1, x: 0, y: 0, rotation: 0 });
  const stateRef = React.useRef(state);
  stateRef.current = state;
  const pointers = React.useRef(new Map<number, { x: number; y: number }>());
  const gesture = React.useRef<{ dist: number; scale: number } | null>(null);
  const [dragging, setDragging] = React.useState(false);

  const fitScale = React.useCallback((): number => {
    const el = surfaceRef.current;
    if (mode === "frame" || !el || !naturalSize || !naturalSize.width || !naturalSize.height) return 1;
    const rotated = Math.abs(stateRef.current.rotation % 180) === 90;
    const w = rotated ? naturalSize.height : naturalSize.width;
    const h = rotated ? naturalSize.width : naturalSize.height;
    // Never enlarge a small image beyond 100% just to fit.
    return Math.min(1, (el.clientWidth - 32) / w, (el.clientHeight - 32) / h);
  }, [mode, naturalSize]);

  const fit = React.useCallback(() => setState((s) => ({ ...s, scale: fitScale(), x: 0, y: 0 })), [fitScale]);

  // New file, or the image size became known: start from "fit".
  React.useEffect(() => {
    setState({ scale: fitScale(), x: 0, y: 0, rotation: 0 });
  }, [resetKey, naturalSize?.width, naturalSize?.height]);

  React.useEffect(() => onScaleChange?.(state.scale), [state.scale]);

  /** Point relative to the surface centre. */
  const centreOffset = (clientX: number, clientY: number) => {
    const r = surfaceRef.current!.getBoundingClientRect();
    return { px: clientX - r.left - r.width / 2, py: clientY - r.top - r.height / 2 };
  };

  React.useImperativeHandle(ref, () => ({
    zoomIn: () => setState((s) => zoomAt(s, s.scale * 1.25, 0, 0)),
    zoomOut: () => setState((s) => zoomAt(s, s.scale / 1.25, 0, 0)),
    fit,
    actualSize: () => setState((s) => ({ ...zoomAt(s, 1, 0, 0), x: 0, y: 0 })),
    rotate: () => setState((s) => ({ ...s, rotation: (s.rotation + 90) % 360, x: 0, y: 0 })),
  }), [fit]);

  // Native (non-passive) wheel listener so the page doesn't scroll while zooming.
  React.useEffect(() => {
    const el = surfaceRef.current;
    if (!el) return;
    const onWheel = (e: WheelEvent) => {
      // Images zoom on any wheel; frames only with Ctrl/⌘ (plain wheel scrolls the document's pages).
      if (mode === "frame" && !e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const { px, py } = centreOffset(e.clientX, e.clientY);
      const factor = e.deltaY < 0 ? WHEEL_STEP : 1 / WHEEL_STEP;
      setState((s) => zoomAt(s, s.scale * factor, px, py));
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, [mode]);

  const panEnabled = mode === "image" || handTool;

  const onPointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    if (!panEnabled || (e.pointerType === "mouse" && e.button !== 0)) return;
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    pointers.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
    if (pointers.current.size === 2) {
      const [a, b] = Array.from(pointers.current.values());
      gesture.current = { dist: Math.hypot(a.x - b.x, a.y - b.y), scale: stateRef.current.scale };
    }
    setDragging(true);
  };

  const onPointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    const prev = pointers.current.get(e.pointerId);
    if (!prev) return;
    pointers.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
    if (pointers.current.size === 2 && gesture.current) {
      // Pinch: zoom around the midpoint of the two fingers.
      const [a, b] = Array.from(pointers.current.values());
      const dist = Math.hypot(a.x - b.x, a.y - b.y);
      const { px, py } = centreOffset((a.x + b.x) / 2, (a.y + b.y) / 2);
      const target = gesture.current.scale * (dist / (gesture.current.dist || dist));
      setState((s) => zoomAt(s, target, px, py));
    } else if (pointers.current.size === 1) {
      const dx = e.clientX - prev.x, dy = e.clientY - prev.y;
      setState((s) => ({ ...s, x: s.x + dx, y: s.y + dy }));
    }
  };

  const endPointer = (e: React.PointerEvent<HTMLDivElement>) => {
    pointers.current.delete(e.pointerId);
    if (pointers.current.size < 2) gesture.current = null;
    if (pointers.current.size === 0) setDragging(false);
  };

  const onDoubleClick = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!panEnabled) return;
    const f = fitScale();
    const { px, py } = centreOffset(e.clientX, e.clientY);
    setState((s) => (s.scale > f * 1.05 ? { ...s, scale: f, x: 0, y: 0 } : zoomAt(s, Math.max(2, f * 2), px, py)));
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    let handled = true;
    switch (e.key) {
      case "+": case "=": setState((s) => zoomAt(s, s.scale * 1.25, 0, 0)); break;
      case "-": case "_": setState((s) => zoomAt(s, s.scale / 1.25, 0, 0)); break;
      case "0": fit(); break;
      case "ArrowLeft": setState((s) => ({ ...s, x: s.x + KEY_PAN_PX })); break;
      case "ArrowRight": setState((s) => ({ ...s, x: s.x - KEY_PAN_PX })); break;
      case "ArrowUp": setState((s) => ({ ...s, y: s.y + KEY_PAN_PX })); break;
      case "ArrowDown": setState((s) => ({ ...s, y: s.y - KEY_PAN_PX })); break;
      default: handled = false;
    }
    if (handled) {
      e.preventDefault();
      e.stopPropagation();
    }
  };

  const transform = `translate(${state.x}px, ${state.y}px) scale(${state.scale}) rotate(${state.rotation}deg)`;

  return (
    <div
      ref={surfaceRef}
      className={`zoomPanSurface ${mode}${panEnabled ? " pannable" : ""}${dragging ? " dragging" : ""}`}
      tabIndex={0}
      role="region"
      aria-label={props.ariaLabel}
      onKeyDown={onKeyDown}
      onPointerDown={mode === "image" ? onPointerDown : undefined}
      onPointerMove={mode === "image" ? onPointerMove : undefined}
      onPointerUp={mode === "image" ? endPointer : undefined}
      onPointerCancel={mode === "image" ? endPointer : undefined}
      onDoubleClick={mode === "image" ? onDoubleClick : undefined}
    >
      <div className="zoomPanContent" style={{ transform }}>
        {props.children}
      </div>
      {mode === "frame" && handTool && (
        // Above the frame, so dragging moves the frame instead of interacting with the document.
        <div
          className="zoomPanHandLayer"
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={endPointer}
          onPointerCancel={endPointer}
          onDoubleClick={onDoubleClick}
        />
      )}
    </div>
  );
});
ZoomPanSurface.displayName = "ZoomPanSurface";
