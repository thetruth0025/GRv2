# TrustedParts logo

`trustedparts-logo.svg` is **deliberately not included**.

TrustedParts require that applications displaying their API data show the words
"Powered by" followed by *their* logo, linked back to trustedparts.com. That logo
is their trademark, so it has to come from them rather than be recreated here.

**To satisfy the requirement:** download the official logo from TrustedParts (ask
`user-requests@trustedparts.com` if it is not on their site) and save it in this
directory as `trustedparts-logo.svg`. Any web format works — change the filename
in `server.py`'s health payload if you use `.png`.

Until then the attribution renders as the words "Powered by TrustedParts.com",
linked and followable. That is visible attribution but **not** the logo their
guidance asks for, so add the file before deploying publicly.

The requirement applies to publicly available applications. Running the app on
localhost for your own use is not covered by it.
