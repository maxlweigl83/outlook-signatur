
function signaturEinfuegen(event) {
  const signatur = `
    <div style="font-family:Arial,sans-serif;font-size:12px;line-height:1.4;">
      <p style="color:#943400;margin:0 0 12px 0;">
        <strong>Markus Mentl-Weigl</strong><br>
        Bürgermeister
      </p>

      <p style="margin:12px 0 16px 0;">
        <img
          src="https://maxlweigl83.github.io/outlook-signatur/image001.png"
          alt="Logo Stadtgemeinde Deutsch-Wagram"
          width="118"
          style="width:118px;height:auto;border:0;"
        >
      </p>

      <p style="color:#174b89;margin:0;">
        Stadtgemeinde Deutsch-Wagram<br>
        2232 Deutsch-Wagram, Bahnhofstraße 1a<br>
        Telefon: +43(0)2247/2209-13<br>
        Fax: +43(0)2247/2209-30<br>
        E-Mail:
        <a href="mailto:buergermeister@deutsch-wagram.gv.at">
          buergermeister@deutsch-wagram.gv.at
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
          console.error("Signaturfehler:", result.error);
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
