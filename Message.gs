function Message() {
    // TODO separate methods
    function nonASCII(str) {
        return '=?UTF-8?B?' + Utilities.base64Encode(str, Utilities.Charset.UTF_8) + '?=';
    }
    this.to = '';
    this.subject = '';
    this.from = '';
    this.body = '';
    this.attachments = [];
    this.toRFC822 = function() {
        // var mt = new MimeType();
        var boundaryHL = 'b314159265359a';
        var separator = '\n';
        var rows = [];
        rows.push('MIME-Version: 1.0');
        rows.push('To: ' + this.to);
        rows.push('Subject: ' + nonASCII(this.subject));
        rows.push('From: ' + this.from);
        rows.push('Content-Type: multipart/mixed; boundary=' + boundaryHL + separator);
        rows.push('--' + boundaryHL);
        rows.push('Content-Type: text/html; charset=UTF-8');
        rows.push('Content-Transfer-Encoding: quoted-printable' + separator);
        rows.push(quotedPrintable.encode(utf8.encode(this.body)));
        for (var i = 0; i < this.attachments.length; i++) {
            rows.push('--' + boundaryHL);
            var fmt = this.attachments[i].getMimeType();
            var fn = this.attachments[i].getName();
            // var ffn = mt.getFullName(fn, fmt);
            // rows.push('Content-Type: ' + fmt + '; charset=UTF-8; name="' + ffn + '"');
            // rows.push('Content-Disposition: attachment; filename="' + ffn + '"');
            rows.push('Content-Type: ' + fmt + '; charset=UTF-8; name="' + fn + '"');
            rows.push('Content-Disposition: attachment; filename="' + fn + '"');
            rows.push('Content-Transfer-Encoding: base64' + separator);
            rows.push(Utilities.base64Encode(this.attachments[i].getBlob().getBytes()));
        }
        rows.push('--' + boundaryHL + '--');
        return rows.join(separator);
    }
    this.createDraft = function() {
        //var me = this;
        var message = {};
        message.raw = Utilities.base64EncodeWebSafe(this.toRFC822());
        if (this.threadId) message.threadId = this.threadId;
        var res = Gmail.Users.Drafts.create({
            'message': message
        }, 'me');

        return res;
    }
}
