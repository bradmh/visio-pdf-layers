using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Layer;
using iText.Kernel.Geom;
using Path = System.IO.Path;
using Visio = Microsoft.Office.Interop.Visio;
using System.Drawing.Printing;
using iText.Kernel.Pdf.Xobject;


namespace pdfExporter
{
    public partial class Ribbon1
    {
        Visio.Application visioApp;
        Visio.Document visioDoc;
        

        private void InitializeVisioApplication()
        {
            try
            {
                visioApp = Marshal.GetActiveObject("Visio.Application") as Visio.Application;
                if (visioApp != null)
                {
                    visioDoc = visioApp.ActiveDocument;
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show("Visio is not currently running or no document is open: " + ex.Message);
            }
        }


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void ExportCurrentPageButton_Click(object sender, RibbonControlEventArgs e)
        {
            InitializeVisioApplication();  // Ensure Visio is initialized

            if (visioApp == null || visioApp.ActiveWindow == null || visioApp.ActiveWindow.Page == null)
            {
                MessageBox.Show("Visio application, window, or page is not available.");
                return;
            }

            string outputPath = GetUserDefinedPath(GetDefaultPath(), GetDefaultFileName());
            if (!string.IsNullOrEmpty(outputPath))
            {
                string pageExportPath = ExportPageWithLayers(visioApp.ActivePage, outputPath);
                CombineLayersIntoOnePDF(pageExportPath, outputPath);
                CleanupTemporaryFiles(GetParentDirectoryPath(pageExportPath));
                MessageBox.Show($"Successfully exported: {outputPath}");
            }
        }


        private void ExportAllPagesButton_Click(object sender, RibbonControlEventArgs e)
        {
            InitializeVisioApplication();  // Ensure Visio is initialized

            if (visioApp == null || visioDoc == null || visioDoc.Pages.Count == 0)
            {
                MessageBox.Show("Visio application is not running or the document does not contain any pages.");
                return;
            }

            string outputPath = GetUserDefinedPath(GetDefaultPath(), GetDefaultFileName());
            string tempPath = "";

            foreach (Visio.Page page in visioDoc.Pages)
            {
                if (page == null) continue;  // Additional check for null page

                string pageExportPath = ExportPageWithLayers(page, outputPath);
                CombineLayersIntoOnePDF(pageExportPath, pageExportPath + ".pdf");
                tempPath = pageExportPath;
            }
            // Combine all exported PDFs into a single document
                
            CombinePDFsIntoSingleDocument(GetParentDirectoryPath(tempPath), outputPath);
            MessageBox.Show($"Successfully exported: {outputPath}");
            CleanupTemporaryFiles(GetParentDirectoryPath(tempPath));
        }

        private string ExportPageWithLayers(Visio.Page page, string outputPath)
        {
            try
            {
                Dictionary<string, bool> originalVisibility = CaptureLayerVisibility(page);
                SetAllLayersVisibility(page, false);

                string baseExportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "visioPDF", SanitizeFileName(visioDoc.Name));
                string pageExportPath = Path.Combine(baseExportPath, SanitizeFileName(page.Name));

                if (!Directory.Exists(pageExportPath))
                {
                    Directory.CreateDirectory(pageExportPath);
                }

                foreach (Visio.Layer layer in page.Layers)
                {
                    if (originalVisibility[layer.Name])
                    {
                        layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = "1"; // Make layer visible
                        string layerFileName = Path.Combine(pageExportPath, $"{SanitizeFileName(layer.Name)}.pdf");
                        SaveAsPDF(layerFileName, page.Index);
                        layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = "0"; // Hide layer again
                    }
                }

                

                RestoreLayerVisibility(page, originalVisibility);

                return pageExportPath; // Return the path where the PDFs are saved or empty string on failure
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export page: " + ex.Message);
                return null;
            }

        }

        public void SaveAsPDF(string outputPath, int pageNumber)
        {
            if (visioDoc != null)
            {

                try
                {
                    // Use ExportAsFixedFormat to export as PDF
                    visioDoc.ExportAsFixedFormat(
                        (VisFixedFormatTypes)VisFixedFormatTypes.visFixedFormatPDF,
                        (string)outputPath,
                        (VisDocExIntent)VisDocExIntent.visDocExIntentScreen,
                        (VisPrintOutRange)PrintRange.Selection,
                        pageNumber,        // From page (ignored if Page Range is 0)
                        pageNumber,        // To page (ignored if Page Range is 0)
                        false,    // Include Background
                        true,     // Include Document Properties
                        false     // Include Structure Tags
                    );
                }
                catch (COMException ex)
                {
                    Console.WriteLine($"Failed to export: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No active document found.");
            }
        }

        private Dictionary<string, bool> CaptureLayerVisibility(Visio.Page page)
        {
            var layerVisibility = new Dictionary<string, bool>();
            foreach (Visio.Layer layer in page.Layers)
            {
                bool isVisible = layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].ResultIU == 1;
                layerVisibility[layer.Name] = isVisible;
            }
            return layerVisibility;
        }

        private void SetAllLayersVisibility(Visio.Page page, bool isVisible)
        {
            foreach (Visio.Layer layer in page.Layers)
            {
                layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = isVisible ? "1" : "0";
            }
        }

        private void RestoreLayerVisibility(Visio.Page page, Dictionary<string, bool> visibilityMap)
        {
            foreach (Visio.Layer layer in page.Layers)
            {
                layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = visibilityMap[layer.Name] ? "1" : "0";
            }
        }

        private string GetUserDefinedPath(string defaultPath, string defaultFileName)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.InitialDirectory = defaultPath;
                saveFileDialog.FileName = defaultFileName;
                saveFileDialog.DefaultExt = ".pdf";
                saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                return saveFileDialog.ShowDialog() == DialogResult.OK ? saveFileDialog.FileName : null;
            }
        }

        private string GetDefaultPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        private string GetDefaultFileName()
        {
            return visioDoc != null ? SanitizeFileName(visioDoc.Name) + ".pdf" : "Document.pdf";
        }

        private string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = Path.GetFileNameWithoutExtension(fileName).Replace(c, '_');
            }
            return fileName;
        }

        private void CombineLayersIntoOnePDF(string directoryPath, string outputPath)
        {
            // Create the output PDF document
            PdfWriter writer = new PdfWriter(outputPath);
            PdfDocument pdf = new PdfDocument(writer);

            // Get all PDF files in this directory
            string[] fileNames = Directory.GetFiles(directoryPath, "*.pdf");

            // Create a new page for this set of layers
            PdfPage page = pdf.AddNewPage(new PageSize(GetFirstPageSize(fileNames[0])));
            PdfCanvas canvas = new PdfCanvas(page, true);

            foreach (string fileName in fileNames)
            {
                PdfDocument srcDoc = new PdfDocument(new PdfReader(fileName));
                PdfFormXObject pageCopy = srcDoc.GetFirstPage().CopyAsFormXObject(pdf);

                // Create a layer (OCG) for the current file
                PdfLayer layer = new PdfLayer(GetDescriptivePartOfFileName(fileName), pdf);
                layer.SetOn(true);  // Set the layer to be visible by default

                // Associate the canvas with the layer
                canvas.BeginLayer(layer);

                // Add the form XObject to the canvas, which contains the source page content
                canvas.AddXObjectAt(pageCopy, 0, 0);

                // End the layer and restore the graphics state
                canvas.EndLayer();

                srcDoc.Close();
            }
            canvas.Release();

            pdf.Close();
        }

        private string GetDescriptivePartOfFileName(string fullPath)
        {
            string filename = System.IO.Path.GetFileNameWithoutExtension(fullPath);
            // Assuming the format is "Page-X_Description"
            int underscoreIndex = filename.IndexOf('_');
            if (underscoreIndex != -1 && underscoreIndex + 1 < filename.Length)
            {
                return filename.Substring(underscoreIndex + 1); // Returns the descriptive part after the underscore
            }
            return filename; // Default to the full filename if no underscore found
        }

        public Rectangle GetFirstPageSize(string pdfFilePath)
        {
            using (PdfDocument pdfDocument = new PdfDocument(new PdfReader(pdfFilePath)))
            {
                PdfPage firstPage = pdfDocument.GetFirstPage();
                Rectangle pageSize = firstPage.GetPageSize();
                return pageSize; // Returns a Rectangle object containing the width and height of the page
            }
        }

        private void CombinePDFsIntoSingleDocument(string baseFolderPath, string outputFileName)
        {
            PdfDocument finalDocument = new PdfDocument(new PdfWriter(outputFileName));
            
            foreach (string file in Directory.GetFiles(baseFolderPath, "*.pdf"))
            {
                PdfDocument sourceDocument = new PdfDocument(new PdfReader(file));
                sourceDocument.CopyPagesTo(1, 1, finalDocument);
                sourceDocument.Close();
            }
            
            finalDocument.Close();
        }

        public void CleanupTemporaryFiles(string sourceFolderPath)
        {
            if (!string.IsNullOrEmpty(sourceFolderPath) && Directory.Exists(sourceFolderPath))
            {
                try
                {
                    // Delete the parent directory and all its contents
                    Directory.Delete(sourceFolderPath, recursive: true);
                    Console.WriteLine($"Deleted parent directory and all contents: {sourceFolderPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete parent directory {sourceFolderPath}: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Parent directory does not exist or the path was the root.");
            }
        }

        private string GetParentDirectoryPath(string directoryPath)
        {
            try
            {
                // Get the directory of the provided path
                string parentDirectoryPath = Path.GetDirectoryName(directoryPath);
                return parentDirectoryPath; // Return the parent directory path 
            }
            catch (Exception ex)
            {
                return directoryPath; // Return the parent directory path 
                Console.WriteLine($"Failed to find parent directory {directoryPath}: {ex.Message}");
            }


        }

        private void InfoButton_Click(object sender, RibbonControlEventArgs e)
        {
            About aboutBox = new About();
            aboutBox.ShowDialog();
        }
    }
}
