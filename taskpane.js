Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        // OneNote-specific initialization code here
    }
});

function applyCustomStyle() {
    // Get the current selection in OneNote
    OneNote.run(function (context) {
        var page = context.application.getActivePage();
        var selectedText = page.contents.getSelectedText();

        // Apply custom style
        selectedText.font.bold = true;
        selectedText.font.color = "red";
        selectedText.font.size = 16;

        return context.sync();
    })
        .catch(function (error) {
            console.log("Error: " + error);
        });
}
