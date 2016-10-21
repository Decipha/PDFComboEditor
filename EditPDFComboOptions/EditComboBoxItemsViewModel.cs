using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quick.MVVM;
using System.Windows;
using System.Collections.ObjectModel;
using System.Windows.Input;
using PDFEditUtil;


namespace EditPDFComboOptions
{
    class EditComboBoxItemsViewModel : ViewModelBase
    {
        public EditComboBoxItemsViewModel()
        {
            this.ComboBoxFieldNames = new List<string>();
            this.ComboOptions       = new List<CboOption>();
            this.WindowTitle        = "Edit Interactive PDF Combo Box Items";


            if (this.InDesignMode)
            {
                this.ComboBoxFieldNames = new List<string>(new[] { "ComboBoxFieldName", "ComboBoxFieldName2" });
                this.ComboBoxFieldName  = "ComboBoxFieldName";

                this.ComboOptions = new List<CboOption>(new[] {
                        new CboOption { Name = "Test Name", Value = "TEST" }
                });
                this.PDFFile = @"C:\Temp\A PDF FILE.pdf";

            }

        }

        /// <summary>
        /// gets or sets the PDF file being 'edited' - if the file exists, it is queried for ComboBox field names.
        /// if there is one ComboBox in the form, that field is selected and the options on that combo are loaded.
        /// </summary>
        public string PDFFile
        {
            get { return GetValue(() => PDFFile); }
            set
            {
                if (SetValue(()=> PDFFile, value))
                {
                    this.WindowTitle = "Editing.... " + value;
                    this.IsBusy = true;
                    try
                    {
                        if (System.IO.File.Exists(value))
                        {
                            this.ComboBoxFieldNames = PDFEditUtil.Forms.GetComboBoxFieldNames(value).ToList();
                            if (this.ComboBoxFieldNames.Count ==1)
                            {
                                this.ComboBoxFieldName = this.ComboBoxFieldNames.First();
                                LoadPDFComboOptions(null);
                            }
                        }
                        
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                    finally
                    {
                        IsBusy = false;
                    }
                }
            }
        }

        /// <summary>
        /// gets or sets the password that is used to unlock the PDF for editing.
        /// </summary>
        public string Pwd
        {
            get { return GetValue(() => Pwd); }
            set { SetValue(() => Pwd, value); }

        }

        /// <summary>
        /// gets or sets the combo-box field being edited.
        /// </summary>
        public string ComboBoxFieldName
        {
            get { return GetValue(() => ComboBoxFieldName); }
            set { SetValue(() => ComboBoxFieldName, value); }
        }

        /// <summary>
        /// list of combo-box field names on the current PDF.
        /// </summary>
        public List<String> ComboBoxFieldNames
        {
            get { return GetValue(() => ComboBoxFieldNames); }
            set { SetValue(() => ComboBoxFieldNames, value); }
        }

        /// <summary>
        /// the list of options that will be set on the selected combo-box when "save" or "save as" are executed.
        /// </summary>
        public List<CboOption> ComboOptions
        {
            get { return GetValue(() => ComboOptions); }
            set { SetValue(() => ComboOptions, value); }
        }

        /// <summary>
        /// loads the combo box items for the selected field from the current pdf.
        /// </summary>
        public ICommand ReadExistingOptions
        {
            get
            {
                return new RelayCommand(() => System.IO.File.Exists(this.PDFFile) && !string.IsNullOrEmpty(this.ComboBoxFieldName), LoadPDFComboOptions);
            }
        }

        /// <summary>
        /// sets the current pdf file.
        /// </summary>
        public ICommand Open
        {
            get { return new RelayCommand(ExecOpenPDF); }
        }

        /// <summary>
        /// saves the list of combo-box options to the selected combo-box field on the current PDF. overwrites the existing file.
        /// </summary>
        public ICommand Save
        {
            get { return new RelayCommand(() => !string.IsNullOrEmpty(this.PDFFile) && !string.IsNullOrEmpty(this.ComboBoxFieldName), ExecSave); }
        }

        /// <summary>
        /// saves the list of combo-box options to the selected combo-box field on the current PDF. prompts for a new file name.
        /// </summary>
        public ICommand SaveAs
        {
            get { return new RelayCommand(() => !string.IsNullOrEmpty(this.PDFFile) && !string.IsNullOrEmpty(this.ComboBoxFieldName), ExecSaveAS); }
        }

        /// <summary>
        /// paste into the list of combo-box items. must be two columns of text data.
        /// </summary>
        public ICommand Paste
        {
            get { return new RelayCommand(()=> Clipboard.ContainsData(DataFormats.CommaSeparatedValue), ExecPaste); }
        }

        /// <summary>
        /// method for "open" command
        /// </summary>
        /// <param name="param"></param>
        protected virtual void ExecOpenPDF(object param)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "PDF Files (*.pdf)|*.pdf";
            dlg.Title = "Open PDF File To Edit";
            var rs = dlg.ShowDialog();
            if (rs.HasValue && rs.Value)
            {
                this.PDFFile = dlg.FileName;
            }
        }

        /// <summary>
        /// method for the Save command. 
        /// replace the options and overwrite the existing file.
        /// </summary>
        /// <param name="param"></param>
        protected virtual void ExecSave(object param)
        {
            if (!string.IsNullOrEmpty(this.PDFFile) && !string.IsNullOrEmpty(this.ComboBoxFieldName))
            {
                this.IsBusy = true;
                try
                {
                    PDFEditUtil.Forms.ReplaceComboOptions(this.PDFFile, this.ComboBoxFieldName, this.ComboOptions, this.Pwd);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                finally
                {
                    IsBusy = false;
                }
            }

        }


        /// <summary>
        /// method for the SaveAs command.
        /// replace the options and prompt for a new file.
        /// </summary>
        /// <param name="param"></param>
        protected virtual void ExecSaveAS(object param)
        {
            if (!string.IsNullOrEmpty(this.PDFFile) && !string.IsNullOrEmpty(this.ComboBoxFieldName))
            {
                if (this.ComboOptions.Count == 0)
                {
                    if (MessageBox.Show($"This will remove all options from the {this.ComboBoxFieldName} combo box. Continue?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                        return;
                }

                var dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.Filter = "PDF Files (*.PDF)|*.PDF";
                dlg.Title  = "Save As";
                var rs = dlg.ShowDialog();
                if (rs.HasValue && rs.Value)
                {
                    this.IsBusy = true;
                    try
                    {
                        PDFEditUtil.Forms.ReplaceComboOptions(this.PDFFile, this.ComboBoxFieldName, this.ComboOptions, this.Pwd, dlg.FileName);
                        

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                    finally
                    {
                        IsBusy = false;
                    }
                }
            }

        }

        /// <summary>
        /// read the list of options in the selected combo-box field for the current pdf.
        /// </summary>
        /// <param name="param"></param>
        protected virtual void LoadPDFComboOptions(object param)
        {
            this.ComboOptions = PDFEditUtil.Forms.GetComboOptions(this.PDFFile, this.ComboBoxFieldName).ToList();
        }

        protected virtual void ExecPaste(object param)
        {

            var data    = GetClipboardData();
            var options = new List<CboOption>(this.ComboOptions);

            foreach (var row in data)
            {

                if (row.Length > 1)
                    options.Add(new CboOption { Name = row[0], Value = row[1] });
            }
            this.ComboOptions = options;


        }
    }
}
