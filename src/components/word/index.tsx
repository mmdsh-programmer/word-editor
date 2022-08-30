import { useEffect } from "react";
import {
  DocumentEditorContainerComponent,
  Toolbar,
} from "@syncfusion/ej2-react-documenteditor";
import TitleBar from "./titleBar";
import { IframeAction, IframeMode } from "../../interface/enum";

DocumentEditorContainerComponent.Inject(Toolbar);
const domains = new Set(["http://localhost:3000"]);

const Word = () => {
  let container: DocumentEditorContainerComponent;

  useEffect(() => {
    window.onbeforeunload = function () {
      return "تغییرات شما ذخیره شود؟";
    };

    container.documentEditor.pageOutline = "#E0E0E0";
    container.documentEditor.resize();
    container.toolbarItems = [
      "New",
      "Open",
      "Separator",
      "Undo",
      "Redo",
      "Separator",
      "Image",
      "Table",
      "Hyperlink",
      "Bookmark",
      "TableOfContents",
      "Separator",
      "Header",
      "Footer",
      "PageSetup",
      "PageNumber",
      "Break",
      "InsertFootnote",
      "InsertEndnote",
      "Separator",
      "Find",
      "Separator",
      "Comments",
      "TrackChanges",
      "LocalClipboard",
      "Separator",
      "FormFields",
      "UpdateFields",
    ];

    new TitleBar(
      document.getElementById("documenteditor_titlebar"),
      container.documentEditor,
      true
    );

    window.addEventListener("message", readEditorData, false);

    return () => {
      window.removeEventListener("message", readEditorData);
    };
  }, []);

  const readEditorData = async (event: MessageEvent) => {
    if (!domains.has(event.origin)) return;
    const { action, key, value } = event.data;
    console.log("iframe says: ", event.data);

    if (action === IframeMode.PREVIEW) {
      container.documentEditor.isReadOnly = value;
      container.showPropertiesPane = !value;
    }

    const exportedDocument = await container.documentEditor.saveAsBlob("Sfdt");
    const reader = new FileReader();
    reader.onload = () => {
      switch (action) {
        case IframeAction.SAVE:
        case IframeAction.FREE_DRAFT:
          event.source!.postMessage(
            {
              action,
              key,
              value: decodeURIComponent(JSON.stringify(reader.result)),
            },
            "*"
          );
          localStorage.setItem(key, JSON.stringify(reader.result));
          break;
        case IframeAction.LOAD:
          if (value) {
            container.documentEditor.open(JSON.parse(value));
          }
          break;
      }
    };

    reader.readAsText(exportedDocument);
  };

  return (
    <div className="control-pane">
      <div className="control-section">
        <div id="documenteditor_titlebar" className="e-de-ctn-title"></div>
        <div id="documenteditor_container_body">
          <DocumentEditorContainerComponent
            id="container"
            ref={(scope) => {
              container = scope!;
            }}
            style={{ display: "block" }}
            height={"100%"}
            locale="en-US"
          />
        </div>
      </div>
    </div>
  );
};

export default Word;
