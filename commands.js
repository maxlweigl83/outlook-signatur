
/**
 * Outlook-Signatur – erste Testversion
 * Stadtgemeinde Deutsch-Wagram
 */

function signaturEinfuegen(event) {
  const signatur = `
    <div style="font-family:Arial,sans-serif;font-size:12px;color:#333;">
      <p>Mit freundlichen Grüßen</p>
      <p>
        <strong>Max Mustermann</strong><br>
        Stadtgemeinde Deutsch-Wagram
      </p>
      <p>
        <a href="https://www.deutsch-wagram.gv.at">
          www.deutsch-wagram.gv.at
        </a>
      </p>
    </div>
  `;

  try {
    Office.context.mailbox.item.body.setSignatureAsync(
      signatur,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Signatur konnte nicht eingefügt werden:", result.error);
        }
        event.completed();
      }
    );
  } catch (error) {
    console.error("Fehler beim Einfügen der Signatur:", error);
    event.completed();
  }
}

Office.actions.associate("signaturEinfuegen", signaturEinfuegen);
