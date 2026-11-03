import * as React from "react";
import { IconButton } from "@fluentui/react/lib/Button";
import { CommandBar, ICommandBarItemProps } from "@fluentui/react/lib/CommandBar";
import { IContextualMenuItem } from "@fluentui/react/lib/ContextualMenu";
import { IFileSystemItem } from "./IFileSystemItem";
import { ItemThumbnail } from "./ItemThumbnail";
import { IStrings } from "./Strings";

export interface IPreviewPaneProps {
  /**
   * Item to preview; a hint is shown if not specified.
   */
  item?: IFileSystemItem;
  /**
   * URL of the preview image of the item.
   */
  imageUrl?: string;
  /**
   * URL of the full preview (Office for the web), shown in a frame instead of the preview image.
   */
  fullPreviewUrl?: string;
  /**
   * Actions of the item, shown in a command bar below the preview.
   */
  actions: IContextualMenuItem[];
  strings: IStrings;
  onClose: () => void;
  /** Opens the selected file in the full-screen viewer (zoom, pan); undefined when not supported. */
  onOpenViewer?: () => void;
}

/**
 * Pane that shows a preview of the selected item over its full height, with the actions of the item below.
 */
export const PreviewPane = (props: IPreviewPaneProps) => {
  const { item, strings } = props;
  // The pane is narrow, so the actions are shown as icons with their names as tooltips.
  const commands: ICommandBarItemProps[] = props.actions.map((action) => ({
    key: action.key,
    text: action.text,
    title: action.text,
    ariaLabel: action.text,
    iconOnly: true,
    iconProps: action.iconProps,
    onClick: () => {
      action.onClick?.();
    },
  }));
  return (
    <div className="previewPane" role="complementary" aria-label={strings.Preview}>
      <div className="previewHeader">
        <span className="previewTitle" title={item?.fullname}>
          {item ? item.fullname : strings.Preview}
        </span>
        {props.onOpenViewer && (
          <IconButton
            iconProps={{ iconName: "FullScreen" }}
            title={strings.OpenViewer}
            ariaLabel={strings.OpenViewer}
            onClick={props.onOpenViewer}
          />
        )}
        <IconButton
          iconProps={{ iconName: "Cancel" }}
          title={strings.Close}
          ariaLabel={strings.Close}
          onClick={props.onClose}
        />
      </div>
      {!item ? (
        <div className="previewHint">{strings.SelectFileToPreview}</div>
      ) : (
        <>
          <div className="previewContent">
            {props.fullPreviewUrl ? (
              <iframe
                // A new frame for each file, so that the previous document is not shown while loading.
                key={props.fullPreviewUrl}
                className="previewFrame"
                src={props.fullPreviewUrl}
                title={item.fullname}
                sandbox="allow-scripts allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-forms allow-downloads"
              />
            ) : (
              <div
                className={`previewImage${props.onOpenViewer ? " clickable" : ""}`}
                onClick={props.onOpenViewer}
                onKeyDown={(e) => {
                  if (props.onOpenViewer && (e.key === "Enter" || e.key === " ")) {
                    e.preventDefault();
                    props.onOpenViewer();
                  }
                }}
                role={props.onOpenViewer ? "button" : undefined}
                tabIndex={props.onOpenViewer ? 0 : undefined}
                title={props.onOpenViewer ? strings.OpenViewer : undefined}
              >
                <ItemThumbnail
                  url={props.imageUrl}
                  className="large"
                  iconProps={{
                    isFolder: item.filetype === "folder",
                    extension: item.filetype,
                    size: "large",
                    className: "tileIcon large",
                    strings,
                  }}
                />
              </div>
            )}
          </div>
          {commands.length > 0 && (
            <CommandBar
              className="previewCommands"
              items={commands}
              ariaLabel={strings.FileActions}
              overflowButtonProps={{ ariaLabel: strings.MoreActions, title: strings.MoreActions }}
            />
          )}
        </>
      )}
    </div>
  );
};
