
(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            }


            $('#Title').click(Title);
            $('#Heading1').click(Heading1);
            $('#Heading2').click(Heading2);
            $('#Heading3').click(Heading3);
            $('#body').click(body);
        });
    };



    function Title() {
        return Word.run(function (context) {
            var range = context.document.getSelection();
            range.paragraphs.getLast().styleBuiltIn = "Title";
            range.font.color = "black";
            range.font.size = 28;
            range.font.name = "EscrowBanner";
            range.paragraphs.getLast().alignment = "Justified";
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().spaceAfter = 0;
            range.paragraphs.getLast().spaceBefore = 0;
            range.load("fuck");
            return context.sync().then(function () { });
        });
    }


    function Heading1() {
        return Word.run(function (context) {
            var range = context.document.getSelection();
            range.font.color = "black";
            range.font.size = 18;
            range.font.name = "Aktiv Grotesk Light";
            range.paragraphs.getLast().alignment = "Justified";
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().spaceAfter = 12;
            range.paragraphs.getLast().spaceBefore = 18;
            range.load("fuck");
            return context.sync().then(function () { });
        });
    }



    function Heading2() {
        return Word.run(function (context) {
            var range = context.document.getSelection();
            range.font.color = "black";
            range.font.size = 14;
            range.font.name = "Aktiv Grotesk Light";
            range.paragraphs.getLast().alignment = "Justified";
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().spaceAfter = 12;
            range.paragraphs.getLast().spaceBefore = 6;
            range.load("fuck");
            return context.sync().then(function () { });
        });
    }



    function Heading3() {
        return Word.run(function (context) {
            var range = context.document.getSelection();
            range.font.color = "black";
            range.font.bold = true;
            range.font.size = 11;
            range.font.name = "Aktiv Grotesk Light";
            range.paragraphs.getLast().alignment = "Justified";
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().spaceAfter = 12;
            range.paragraphs.getLast().spaceBefore = 6;
            range.load("fuck");
            return context.sync().then(function () { });
        });
    }


    function body() {
        return Word.run(function (context) {
            var range = context.document.getSelection();
            range.font.color = "black";
            range.font.bold = true;
            range.font.size = 10;
            range.font.name = "Aktiv Grotesk Light";
            range.paragraphs.getLast().alignment = "Justified";
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().lineSpacing = 12;
            range.paragraphs.getLast().spaceAfter = 0;
            range.paragraphs.getLast().spaceBefore = 6;
            range.load("fuck");
            return context.sync().then(function () { });
        });
    }
})();