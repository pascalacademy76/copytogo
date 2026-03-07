/* global document, Office, PowerPoint */

let copiedProperties = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    document.getElementById("copy-btn").onclick = copyProperties;
    document.getElementById("paste-btn").onclick = pasteProperties;
  }
});

async function copyProperties() {
  try {
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      const shapeCount = selectedShapes.getCount();
      await context.sync();

      if (shapeCount.value !== 1) {
        document.getElementById("status").innerText = "⚠️ Selecteer een afbeelding.";
        return;
      }

      const shape = selectedShapes.getItemAt(0);
      shape.load("left,top,width,height");
      await context.sync();

      copiedProperties = {
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height
      };

      document.getElementById("status").innerText = "✅ Eigenschappen gekopieerd!";
    });
  } catch (error) {
    document.getElementById("status").innerText = "❌ Fout: " + error.message;
  }
}

async function pasteProperties() {
  try {
    if (!copiedProperties) {
      document.getElementById("status").innerText = "⚠️ Kopieer eerst eigenschappen.";
      return;
    }

    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      const shapeCount = selectedShapes.getCount();
      await context.sync();

      if (shapeCount.value !== 1) {
        document.getElementById("status").innerText = "⚠️ Selecteer een afbeelding.";
        return;
      }

      const shape = selectedShapes.getItemAt(0);
      shape.left = copiedProperties.left;
      shape.top = copiedProperties.top;
      shape.width = copiedProperties.width;
      shape.height = copiedProperties.height;

      await context.sync();
      document.getElementById("status").innerText = "✅ Eigenschappen geplakt!";
    });
  } catch (error) {
    document.getElementById("status").innerText = "❌ Fout: " + error.message;
  }
}