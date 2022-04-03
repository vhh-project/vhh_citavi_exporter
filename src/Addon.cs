
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Controls;
using SwissAcademic.Drawing;
using System;
using System.Windows.Forms;

namespace macroAddon
{
    public class Addon : CitaviAddOn
    {
        #region Properties
        public override AddOnHostingForm HostingForm
        {
            get { return AddOnHostingForm.MainForm; }
        }

        public string MB_SETFOREGROUND { get; private set; }

        #endregion

        #region Methods

        protected override void OnApplicationIdle(Form form)
        {
            base.OnApplicationIdle(form);
        }

        protected override void OnBeforePerformingCommand(BeforePerformingCommandEventArgs e)
        {
            switch (e.Key)
            {
                case "ExportProjectToJson":
                    {
                        e.Handled = true;
                        FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

                        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                        {
                            var prefix = folderBrowserDialog.SelectedPath;
                        
                            String name = Program.ActiveProjectShell.Project.Name;
                            prefix = prefix + "\\CitaviExport_"+ name;
                            System.IO.Directory.CreateDirectory(prefix);

                            Categories.Export(prefix);
                            Groups.Export(prefix);
                            Keywords.Export(prefix);
                            KnowledgeItems.Export(prefix);
                            Papers.Export(prefix);
                            
                            System.IO.Directory.CreateDirectory(prefix + "\\Files");
                            String log = Pdfs.Export(prefix);

       

						    System.Windows.Forms.MessageBox.Show("Done exporting.","Finished");
                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(prefix+"\\log.txt", true))
                            {
                                file.WriteLine(log);
                            }
                        }
                    }
                    break;
            }
            base.OnBeforePerformingCommand(e);
        }

        protected override void OnChangingColorScheme(Form form, ColorScheme colorScheme)
        {
            base.OnChangingColorScheme(form, colorScheme);
        }

        protected override void OnHostingFormLoaded(Form form)
        {
           
            MainForm mainForm = (MainForm)form;
            mainForm.GetMainCommandbarManager()
                .GetReferenceEditorCommandbar(MainFormReferenceEditorCommandbarId.Menu)
                .GetCommandbarMenu(MainFormReferenceEditorCommandbarMenuId.Tools)
                .AddCommandbarButton("ExportProjectToJson", "Export Project to JSON");
            base.OnHostingFormLoaded(form);
        }

        protected override void OnLocalizing(Form form)
        {
            base.OnLocalizing(form);
        }

        #endregion
    }
}