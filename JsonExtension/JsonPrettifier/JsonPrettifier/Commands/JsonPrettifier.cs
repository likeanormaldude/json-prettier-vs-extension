using System.Linq;
using EnvDTE;
using EnvDTE80;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonPrettifier;

[Command(PackageIds.MyCommand)]
internal sealed class JsonPrettifier : BaseCommand<JsonPrettifier>
{
    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        DTE2 dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
        Document activeDoc = dte?.ActiveDocument;

        if (activeDoc == null)
        {
            await VS.MessageBox.ShowAsync("No active document!");
            return;
        }

        TextDocument textDoc = activeDoc.Object() as TextDocument;
        EditPoint startPoint = textDoc?.StartPoint.CreateEditPoint();

        if (textDoc == null || startPoint == null)
        {
            await VS.MessageBox.ShowAsync("Cannot access text in the current document.");
            return;
        }

        string text = startPoint.GetText(textDoc.EndPoint);

        try
        {
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var prettifiedLines = lines.Select(line =>
            {
                var parsed = JToken.Parse(line);
                return parsed.ToString(Formatting.Indented);
            });

            string pretty = string.Join(Environment.NewLine, prettifiedLines);
            startPoint.ReplaceText(textDoc.EndPoint, pretty, 0);
        }
        catch (JsonReaderException ex)
        {
            await VS.MessageBox.ShowAsync("Invalid JSON:\n" + ex.Message);
        }
        catch (Exception ex)
        {
            await VS.MessageBox.ShowAsync("Error:\n" + ex.Message);
        }
    }
}
