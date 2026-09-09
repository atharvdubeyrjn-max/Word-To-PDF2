using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordToPDF;

public sealed class MainForm : Form
{
    private readonly Button selectButton = new();
    private readonly Button convertButton = new();
    private readonly Button copyButton = new();
    private readonly Button openButton = new();
    private readonly Button anotherButton = new();
    private readonly Label fileLabel = new();
    private readonly Label statusLabel = new();
    private string? selectedWordFile;
    private string? outputPdf;

    public MainForm()
    {
        Text = "Word to PDF";
        Width = 560;
        Height = 390;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        AllowDrop = true;

        var title = new Label
        {
            Text = "Word to PDF",
            Font = new System.Drawing.Font("Segoe UI", 20, System.Drawing.FontStyle.Bold),
            AutoSize = true,
            Left = 32,
            Top = 25
        };

        var subtitle = new Label
        {
            Text = "Convert Word documents to PDF — completely offline.",
            AutoSize = true,
            Left = 35,
            Top = 67
        };

        fileLabel.Text = "No Word file selected";
        fileLabel.Left = 35;
        fileLabel.Top = 105;
        fileLabel.Width = 475;
        fileLabel.AutoEllipsis = true;

        selectButton.Text = "Select Word File";
        selectButton.SetBounds(35, 140, 150, 42);
        selectButton.Click += (_, _) => SelectFile();

        convertButton.Text = "Convert to PDF";
        convertButton.SetBounds(200, 140, 150, 42);
        convertButton.Enabled = false;
        convertButton.Click += (_, _) => ConvertToPdf();

        copyButton.Text = "Copy PDF";
        copyButton.SetBounds(365, 140, 110, 42);
        copyButton.Enabled = false;
        copyButton.Click += (_, _) => CopyPdf();

        openButton.Text = "Open PDF";
        openButton.SetBounds(35, 195, 150, 38);
        openButton.Enabled = false;
        openButton.Click += (_, _) => OpenPdf();

        anotherButton.Text = "Convert Another";
        anotherButton.SetBounds(200, 195, 150, 38);
        anotherButton.Enabled = false;
        anotherButton.Click += (_, _) => ResetApp();

        var dropLabel = new Label
        {
            Text = "Tip: You can also drag and drop a .doc or .docx file here.",
            AutoSize = true,
            Left = 35,
            Top = 255
        };

        statusLabel.Text = "Ready";
        statusLabel.Left = 35;
        statusLabel.Top = 300;
        statusLabel.Width = 475;
        statusLabel.AutoEllipsis = true;

        Controls.AddRange(new Control[]
        {
            title, subtitle, fileLabel, selectButton, convertButton,
            copyButton, openButton, anotherButton, dropLabel, statusLabel
        });

        DragEnter += MainForm_DragEnter;
        DragDrop += MainForm_DragDrop;
    }

    private void MainForm_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            e.Effect = DragDropEffects.Copy;
    }

    private void MainForm_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            SetSelectedFile(files[0]);
    }

    private void SelectFile()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Word Documents (*.doc;*.docx)|*.doc;*.docx",
            Title = "Select a Word document"
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
            SetSelectedFile(dialog.FileName);
    }

    private void SetSelectedFile(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext != ".doc" && ext != ".docx")
        {
            MessageBox.Show(this, "Please select a .doc or .docx Word document.",
                "Invalid file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        selectedWordFile = path;
        outputPdf = Path.Combine(Path.GetDirectoryName(path)!,
            Path.GetFileNameWithoutExtension(path) + ".pdf");

        fileLabel.Text = Path.GetFileName(path);
        convertButton.Enabled = true;
        copyButton.Enabled = false;
        openButton.Enabled = false;
        anotherButton.Enabled = false;
        statusLabel.Text = "Ready to convert.";
    }

    private void ConvertToPdf()
    {
        if (string.IsNullOrWhiteSpace(selectedWordFile) || !File.Exists(selectedWordFile))
            return;

        if (File.Exists(outputPdf))
        {
            var result = MessageBox.Show(this,
                $"A PDF named \"{Path.GetFileName(outputPdf)}\" already exists.\n\nReplace it?",
                "PDF already exists", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;
        }

        dynamic? word = null;
        dynamic? document = null;

        try
        {
            statusLabel.Text = "Converting...";
            Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new InvalidOperationException(
                    "Microsoft Word was not found on this computer. Please install Microsoft Word first.");

            word = Activator.CreateInstance(wordType);
            word.Visible = false;
            word.DisplayAlerts = 0;

            document = word.Documents.Open(
                selectedWordFile,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            // Word WdExportFormat = 17 (PDF), WdExportOptimizeFor = 0 (Print),
            // WdExportRange = 0 (AllDocument), WdExportItem = 0 (DocumentContent),
            // WdExportCreateBookmarks = 1 (HeadingBookmarks).
            document.ExportAsFixedFormat(
                outputPdf, 17, false, 0, 0, 1, 1, 0, true, false, 1, true, true, false);

            statusLabel.Text = "PDF created successfully.";
            copyButton.Enabled = true;
            openButton.Enabled = true;
            anotherButton.Enabled = true;

            MessageBox.Show(this,
                $"Done!\n\nSaved as:\n{outputPdf}\n\nClick \"Copy PDF\" to paste it into WhatsApp.",
                "Conversion complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (COMException)
        {
            MessageBox.Show(this,
                "Microsoft Word could not convert this document. Make sure Word is installed and the document is not damaged.",
                "Conversion failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            statusLabel.Text = "Conversion failed.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Conversion failed",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            statusLabel.Text = "Conversion failed.";
        }
        finally
        {
            try { document?.Close(false); } catch { }
            try { word?.Quit(); } catch { }

            if (document != null) Marshal.FinalReleaseComObject(document);
            if (word != null) Marshal.FinalReleaseComObject(word);

            Cursor = Cursors.Default;
        }
    }

    private void CopyPdf()
    {
        if (string.IsNullOrWhiteSpace(outputPdf) || !File.Exists(outputPdf))
            return;

        var files = new StringCollection();
        files.Add(outputPdf);
        Clipboard.SetFileDropList(files);
        statusLabel.Text = "PDF copied. Paste it into WhatsApp with Ctrl+V.";
        MessageBox.Show(this,
            "PDF copied to the clipboard.\n\nNow open WhatsApp and press Ctrl+V.",
            "PDF copied", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void OpenPdf()
    {
        if (!string.IsNullOrWhiteSpace(outputPdf) && File.Exists(outputPdf))
            Process.Start(new ProcessStartInfo(outputPdf) { UseShellExecute = true });
    }

    private void ResetApp()
    {
        selectedWordFile = null;
        outputPdf = null;
        fileLabel.Text = "No Word file selected";
        statusLabel.Text = "Ready";
        convertButton.Enabled = false;
        copyButton.Enabled = false;
        openButton.Enabled = false;
        anotherButton.Enabled = false;
    }
}
