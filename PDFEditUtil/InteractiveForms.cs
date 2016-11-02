using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Pdf.InteractiveFeatures.Forms;
using Aspose.Pdf;



namespace PDFEditUtil
{
    /// <summary>
    /// entity for returning combo-box option values (abstracts away from using Aspose objects)
    /// </summary>
    public class CboOption
    {
        /// <summary>
        /// this is what is shown on the PDF
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// this is the "export value"
        /// </summary>
        public string Value { get; set; }

    }

    /// <summary>
    /// static methods for querying and modifying PDF interactive forms.
    /// </summary>
	public class Forms
	{
        /// <summary>
        /// static initializer to setup the Aspose licence.
        /// </summary>
        static Forms()
        {
            // Create a PDF license object
            License license = new License();
            // Instantiate license file
            using (var ms = new System.IO.MemoryStream(Properties.Resources.Aspose_Total))
                license.SetLicense(ms);

            // Set the value to indicate that license will be embedded in the application
            license.Embedded = true;

        }

        /// <summary>
        /// yields an enumeration of all field names in the specified interactive PDF.
        /// </summary>
        /// <param name="fileName">the PDF file to examine</param>
        /// <returns>field names</returns>
        public static IEnumerable<string> GetFieldNames(string fileName)
        {
            using (var pdf = new Aspose.Pdf.Document(fileName))
            {
                foreach (var fld in pdf.Form.Fields)
                    yield return fld.FullName;
            }
        }

        /// <summary>
        /// yields an enumeration of field-names that are ComboBoxes in the specified interactive PDF.
        /// </summary>
        /// <param name="fileName">the PDF file to examine</param>
        /// <returns>combo-box field namdes (field.FullName)</returns>
        public static IEnumerable<string> GetComboBoxFieldNames(string fileName)
        {
            using (var pdf = new Document(fileName))
            {
                foreach (var fld in pdf.Form.Fields)
                    if (fld is Aspose.Pdf.InteractiveFeatures.Forms.ComboBoxField)
                        yield return fld.FullName;
            }
        }


        /// <summary>
        /// returns a list of options for the specified combo box in the specified PDF file.
        /// </summary>
        /// <param name="pdfFileName">the PDF file</param>
        /// <param name="comboBoxFieldName">the combo box field name (use <see cref="GetComboBoxFieldNames(string)"/> to query appropriate field names)</param>
        /// <returns>
        /// list of options (items) available for the specified combo box.
        /// </returns>
        public static IEnumerable<CboOption> GetComboOptions(string pdfFileName, string comboBoxFieldName)
        {
            // open the pdf
            using (var pdf = new Document(pdfFileName))
            {
                // find the field
                var field = pdf.Form.Fields.FirstOrDefault((f) => f.FullName.Equals(comboBoxFieldName));
                if (field != null)
                {
                    // check the field is a combo box
                    if (field is ComboBoxField)
                    {
                        // cast to a combo box
                        var combo = field as ComboBoxField;
                        foreach (Option option in combo.Options)
                        {
                            // yield each option
                            yield return new CboOption { Name = option.Name, Value = option.Value };
                        }

                    }

                }
            }


        }

        /// <summary>
        /// replaces the list of options for the specified combo box.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fieldName"></param>
        /// <param name="options"></param>
        /// <param name="password"></param>
        /// <param name="outputFileName"></param>
        public static void ReplaceComboOptions(string fileName, string fieldName, IEnumerable<CboOption> options, string password = null, string outputFileName = null)
        {
            using (var pdf = (string.IsNullOrEmpty(password) ? new Document(fileName) : new Document(fileName, password)))
            {
                var field = pdf.Form.Fields.FirstOrDefault((f) => f.FullName.Equals(fieldName));
                if (field != null)
                {
                    if (field is ComboBoxField)
                    {
                        var combo = field as ComboBoxField;
                        var optionNames = (from Option o in combo.Options select o.Name).ToArray();
                        foreach (var optionName in optionNames)
                        {
                            combo.DeleteOption(optionName);
                        }
                        foreach (var option in options.OrderBy((o)=>o.Name))
                        {
                            combo.AddOption(option.Value, option.Name);
                        }
                        if (string.IsNullOrEmpty(outputFileName))
                            pdf.Save(fileName);
                        else
                            pdf.Save(outputFileName);
                    }

                }
            }
        }

        /// <summary>
        /// adds to the list of options for the specified combo box.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fieldName"></param>
        /// <param name="password"></param>
        /// <param name="options"></param>
        /// <param name="outputFileName"></param>
        public static void AddComboOptions(string fileName, string fieldName, string password, IEnumerable<CboOption> options, string outputFileName)
        {
            using (var pdf = (string.IsNullOrEmpty(password) ? new Document(fileName) : new Document(fileName, password)))
            {
                var field = pdf.Form.Fields.FirstOrDefault((f) => f.FullName.Equals(fieldName));
                if (field != null)
                {
                    if (field is ComboBoxField)
                    {
                        var combo = field as ComboBoxField;
                        foreach (var option in options.OrderBy((o) => o.Name))
                        {
                            combo.AddOption(option.Value, option.Name);
                        }
                        if (string.IsNullOrEmpty(outputFileName))
                            pdf.Save(fileName);
                        else
                            pdf.Save(outputFileName);
                    }

                }
            }
        }

    }
}
