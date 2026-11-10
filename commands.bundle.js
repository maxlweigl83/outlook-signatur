
Office.actions.associate("signaturEinfuegen", function (event) {
  const signatur = `
    <div style="font-family:Calibri,Arial,sans-serif;font-size:12pt;line-height:normal;">
      <p style="margin:0;color:#833C0B;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        <strong>Markus Mentl-Weigl</strong>
      </p>
      <p style="margin:0;color:#833C0B;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        Bürgermeister
      </p>

      <p style="margin:12pt 0 0 0;">
        <img
          src="https://maxlweigl83.github.io/outlook-signatur/image001.png"
          alt="Logo Stadtgemeinde Deutsch-Wagram"
          width="163"
          height="85"
          style="width:163px;height:85px;border:0;"
        >
      </p>

      <p style="margin:12pt 0 0 0;color:#1F497D;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        Stadtgemeinde Deutsch-Wagram
      </p>
      <p style="margin:0;color:#1F497D;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        2232 Deutsch-Wagram, Bahnhofstraße 1a
      </p>
      <p style="margin:0;color:#1F497D;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        Telefon: +43(0)2247/2209-13
      </p>
      <p style="margin:0;color:#1F497D;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        Fax: +43(0)2247/2209-30
      </p>
      <p style="margin:0;color:#1F497D;font-family:Calibri,Arial,sans-serif;font-size:12pt;">
        E-Mail:
        <a href="mailto:buergermeister@deutsch-wagram.gv.at">buergermeister@deutsch-wagram.gv.at</a>
      </p>

      <table role="presentation" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td height="30" style="height:30px;font-size:1px;line-height:30px;">&nbsp;</td>
        </tr>
        <tr>
          <td>
            <img
              src="https://maxlweigl83.github.io/outlook-signatur/Aktueller-Banner.jpg"
              alt="Aktuelles Kampagnenbanner"
              width="600"
              style="display:block;width:600px;max-width:100%;height:auto;border:0;"
            >
          </td>
        </tr>
      </table>
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
});
