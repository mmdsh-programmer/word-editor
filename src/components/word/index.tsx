import { useEffect, useRef, useState } from "react";
import {
  DocumentEditorContainerComponent,
  Toolbar,
} from "@syncfusion/ej2-react-documenteditor";
import TitleBar from "./titleBar";
import { IframeAction } from "../../interface/enum";

DocumentEditorContainerComponent.Inject(Toolbar);
const domains = new Set(["http://localhost:3000"]);

const Word = () => {
  let container!: DocumentEditorContainerComponent;

  // const loadSampleData = (data: string) => {
  //   console.log("main app json", JSON.parse(data));
  //   container.documentEditor.open(JSON.parse(data));
  // };

  useEffect(() => {
    window.onbeforeunload = function () {
      return "تغییرات شما ذخیره شود؟";
    };

    container.documentEditor.pageOutline = "#E0E0E0";
    container.documentEditor.acceptTab = true;
    container.documentEditor.resize();
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
    console.log(event.data);

    const exportedDocument = await container.documentEditor.saveAsBlob("Sfdt");
    const reader = new FileReader();
    reader.onload = () => {
      switch (action) {
        case IframeAction.SAVE:
          event.source!.postMessage(
            {
              action,
              key,
              value: JSON.stringify(reader.result),
            },
            "*"
          );
          localStorage.setItem(key, JSON.stringify(reader.result));
          break;
        case IframeAction.LOAD:
          debugger;
          console.log("main app", value);
          break;
      }
    };

    reader.readAsText(exportedDocument, "ISO-8859-1");
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
            enableToolbar={true}
            locale="en-US"
          />
        </div>
      </div>
    </div>
  );
};

export default Word;
