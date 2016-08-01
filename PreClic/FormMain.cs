#define DEBUG
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web;
using System.Windows.Forms;
using System.Xml.Linq;
using PreClic.Helpers;
using PreClic.Properties;
using System.Net.NetworkInformation;

namespace PreClic
{
  public partial class FormMain : Form
  {
    public FormMain()
    {
      InitializeComponent();
    }

    public readonly Dictionary<string, string> _languageDicoEn = new Dictionary<string, string>();
    public readonly Dictionary<string, string> _languageDicoFr = new Dictionary<string, string>();
    private string _currentLanguage = "english";
    private ConfigurationOptions _configurationOptions = new ConfigurationOptions();
    private DataTable dataTableCsvFile = new DataTable();
    private string workingDirectory;

    private void QuitToolStripMenuItem_Click(object sender, EventArgs e)
    {
      SaveWindowValue();
      Application.Exit();
    }

    private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
    {
      AboutBoxApplication aboutBoxApplication = new AboutBoxApplication();
      aboutBoxApplication.ShowDialog();
    }

    private void DisplayTitle()
    {
      Assembly assembly = Assembly.GetExecutingAssembly();
      FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
      Text += string.Format(" V{0}.{1}.{2}.{3}", fvi.FileMajorPart, fvi.FileMinorPart, fvi.FileBuildPart, fvi.FilePrivatePart);
    }

    private void FormMain_Load(object sender, EventArgs e)
    {
      LoadSettingsAtStartup();
    }

    private void LoadSettingsAtStartup()
    {
      DisplayTitle();
      GetWindowValue();
      LoadLanguages();
      SetLanguage(Settings.Default.LastLanguageUsed);
      workingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
    }

    private void LoadConfigurationOptions()
    {
      _configurationOptions.Option1Name = Settings.Default.Option1Name;
      _configurationOptions.Option2Name = Settings.Default.Option2Name;
    }

    private void SaveConfigurationOptions()
    {
      _configurationOptions.Option1Name = Settings.Default.Option1Name;
      _configurationOptions.Option2Name = Settings.Default.Option2Name;
    }

    private void LoadLanguages()
    {
      if (!File.Exists(Settings.Default.LanguageFileName))
      {
        CreateLanguageFile();
      }

      // read the translation file and feed the language
      XDocument xDoc;
      try
      {
        xDoc = XDocument.Load(Settings.Default.LanguageFileName);
      }
      catch (Exception exception)
      {
        MessageBox.Show(Resources.Error_while_loading_the + Punctuation.OneSpace +
          Settings.Default.LanguageFileName + Punctuation.OneSpace + Resources.XML_file +
          Punctuation.OneSpace + exception.Message);
        CreateLanguageFile();
        return;
      }
      var result = from node in xDoc.Descendants("term")
                   where node.HasElements
                   let xElementName = node.Element("name")
                   where xElementName != null
                   let xElementEnglish = node.Element("englishValue")
                   where xElementEnglish != null
                   let xElementFrench = node.Element("frenchValue")
                   where xElementFrench != null
                   select new
                   {
                     name = xElementName.Value,
                     englishValue = xElementEnglish.Value,
                     frenchValue = xElementFrench.Value
                   };
      foreach (var i in result)
      {
        if (!_languageDicoEn.ContainsKey(i.name))
        {
          _languageDicoEn.Add(i.name, i.englishValue);
        }
#if DEBUG
        else
        {
          MessageBox.Show(Resources.Your_XML_file_has_duplicate_like + Punctuation.Colon +
            Punctuation.OneSpace + i.name);
        }
#endif
        if (!_languageDicoFr.ContainsKey(i.name))
        {
          _languageDicoFr.Add(i.name, i.frenchValue);
        }
#if DEBUG
        else
        {
          MessageBox.Show(Resources.Your_XML_file_has_duplicate_like + Punctuation.Colon +
            Punctuation.OneSpace + i.name);
        }
#endif
      }
    }

    private static void CreateLanguageFile()
    {
      List<string> minimumVersion = new List<string>
      {
        "<?xml version=\"1.0\" encoding=\"utf-8\" ?>",
        "<terms>",
         "<term>",
        "<name>MenuFile</name>",
        "<englishValue>File</englishValue>",
        "<frenchValue>Fichier</frenchValue>",
        "</term>",
        "<term>",
        "<name>MenuFileNew</name>",
        "<englishValue>New</englishValue>",
        "<frenchValue>Nouveau</frenchValue>",
        "</term>",
        "<term>",
        "<name>MenuFileOpen</name>",
        "<englishValue>Open</englishValue>",
        "<frenchValue>Ouvrir</frenchValue>",
        "</term>",
        "<term>",
        "<name>MenuFileSave</name>",
        "<englishValue>Save</englishValue>",
        "<frenchValue>Enregistrer</frenchValue>",
        "</term>",
        "<term>",
        "<name>MenuFileSaveAs</name>",
        "<englishValue>Save as ...</englishValue>",
        "<frenchValue>Enregistrer sous ...</frenchValue>",
        "</term>",
        "<term>",
        "<name>MenuFilePrint</name>",
        "<englishValue>Print ...</englishValue>",
        "<frenchValue>Imprimer ...</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenufilePageSetup</name>",
          "<englishValue>Page setup</englishValue>",
          "<frenchValue>Aperçu avant impression</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenufileQuit</name>",
          "<englishValue>Quit</englishValue>",
          "<frenchValue>Quitter</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEdit</name>",
          "<englishValue>Edit</englishValue>",
          "<frenchValue>Edition</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEditCancel</name>",
          "<englishValue>Cancel</englishValue>",
          "<frenchValue>Annuler</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEditRedo</name>",
          "<englishValue>Redo</englishValue>",
          "<frenchValue>Rétablir</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEditCut</name>",
          "<englishValue>Cut</englishValue>",
          "<frenchValue>Couper</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEditCopy</name>",
          "<englishValue>Copy</englishValue>",
          "<frenchValue>Copier</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEditPaste</name>",
          "<englishValue>Paste</englishValue>",
          "<frenchValue>Coller</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuEditSelectAll</name>",
          "<englishValue>Select All</englishValue>",
          "<frenchValue>Sélectionner tout</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuTools</name>",
          "<englishValue>Tools</englishValue>",
          "<frenchValue>Outils</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuToolsCustomize</name>",
          "<englishValue>Customize ...</englishValue>",
          "<frenchValue>Personaliser ...</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuToolsOptions</name>",
          "<englishValue>Options</englishValue>",
          "<frenchValue>Options</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuLanguage</name>",
          "<englishValue>Language</englishValue>",
          "<frenchValue>Langage</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuLanguageEnglish</name>",
          "<englishValue>English</englishValue>",
          "<frenchValue>Anglais</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuLanguageFrench</name>",
          "<englishValue>French</englishValue>",
          "<frenchValue>Français</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuHelp</name>",
          "<englishValue>Help</englishValue>",
          "<frenchValue>Aide</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuHelpSummary</name>",
          "<englishValue>Summary</englishValue>",
          "<frenchValue>Sommaire</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuHelpIndex</name>",
          "<englishValue>Index</englishValue>",
          "<frenchValue>Index</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuHelpSearch</name>",
          "<englishValue>Search</englishValue>",
          "<frenchValue>Rechercher</frenchValue>",
        "</term>",
        "<term>",
          "<name>MenuHelpAbout</name>",
          "<englishValue>About</englishValue>",
          "<frenchValue>A propos de ...</frenchValue>",
        "</term>",
        "</terms>"
      };
      StreamWriter sw = new StreamWriter(Settings.Default.LanguageFileName);
      foreach (string item in minimumVersion)
      {
        sw.WriteLine(item);
      }

      sw.Close();
    }

    private void GetWindowValue()
    {
      Width = Settings.Default.WindowWidth;
      Height = Settings.Default.WindowHeight;
      Top = Settings.Default.WindowTop < 0 ? 0 : Settings.Default.WindowTop;
      Left = Settings.Default.WindowLeft < 0 ? 0 : Settings.Default.WindowLeft;
      SetDisplayOption(Settings.Default.DisplayToolStripMenuItem);
      LoadConfigurationOptions();
      textBoxCsvFilePath.Text = Settings.Default.textBoxCsvFilePath;
      textBoxDecryptFilePath.Text = Settings.Default.textBoxDecryptFilePath;
      tabControlMain.SelectedIndex = Settings.Default.MostRecentTabUsed;
      textBoxWebBrowserUrl.Text = Settings.Default.textBoxWebBrowserUrl;
    }

    private void SaveWindowValue()
    {
      Settings.Default.WindowHeight = Height;
      Settings.Default.WindowWidth = Width;
      Settings.Default.WindowLeft = Left;
      Settings.Default.WindowTop = Top;
      Settings.Default.LastLanguageUsed = frenchToolStripMenuItem.Checked ? "French" : "English";
      Settings.Default.DisplayToolStripMenuItem = GetDisplayOption();
      SaveConfigurationOptions();
      Settings.Default.textBoxCsvFilePath = textBoxCsvFilePath.Text;
      Settings.Default.textBoxDecryptFilePath = textBoxDecryptFilePath.Text;
      Settings.Default.MostRecentTabUsed = tabControlMain.SelectedIndex;
      if (dataGridViewMain.Columns.Count > 0)
      {
        Settings.Default.DgvSizeColum0 = dataGridViewMain.Columns[0].Width;
        Settings.Default.DgvSizeColum1 = dataGridViewMain.Columns[1].Width;
        Settings.Default.DgvSizeColum2 = dataGridViewMain.Columns[2].Width;
        Settings.Default.DgvSizeColum3 = dataGridViewMain.Columns[3].Width;
        Settings.Default.DgvSizeColum4 = dataGridViewMain.Columns[4].Width;
        Settings.Default.DgvSizeColum5 = dataGridViewMain.Columns[5].Width;
        Settings.Default.DgvSizeColum6 = dataGridViewMain.Columns[6].Width;
        Settings.Default.DgvSizeColum7 = dataGridViewMain.Columns[7].Width;
        Settings.Default.DgvSizeColum8 = dataGridViewMain.Columns[8].Width;
        Settings.Default.DgvSizeColum9 = dataGridViewMain.Columns[9].Width;
        Settings.Default.DgvSizeColum10 = dataGridViewMain.Columns[10].Width;
        Settings.Default.DgvSizeColum11 = dataGridViewMain.Columns[11].Width;
      }

      Settings.Default.textBoxWebBrowserUrl = textBoxWebBrowserUrl.Text;
      Settings.Default.Save();
    }

    private string GetDisplayOption()
    {
      if (SmallToolStripMenuItem.Checked)
      {
        return "Small";
      }

      if (MediumToolStripMenuItem.Checked)
      {
        return "Medium";
      }

      return LargeToolStripMenuItem.Checked ? "Large" : string.Empty;
    }

    private void SetDisplayOption(string option)
    {
      UncheckAllOptions();
      switch (option.ToLower())
      {
        case "small":
          SmallToolStripMenuItem.Checked = true;
          break;
        case "medium":
          MediumToolStripMenuItem.Checked = true;
          break;
        case "large":
          LargeToolStripMenuItem.Checked = true;
          break;
        default:
          SmallToolStripMenuItem.Checked = true;
          break;
      }
    }

    private void UncheckAllOptions()
    {
      SmallToolStripMenuItem.Checked = false;
      MediumToolStripMenuItem.Checked = false;
      LargeToolStripMenuItem.Checked = false;
    }

    private void FormMainFormClosing(object sender, FormClosingEventArgs e)
    {
      SaveWindowValue();
    }

    private void frenchToolStripMenuItem_Click(object sender, EventArgs e)
    {
      _currentLanguage = Language.French.ToString();
      SetLanguage(Language.French.ToString());
      AdjustAllControls();
    }

    private void englishToolStripMenuItem_Click(object sender, EventArgs e)
    {
      _currentLanguage = Language.English.ToString();
      SetLanguage(Language.English.ToString());
      AdjustAllControls();
    }

    private void SetLanguage(string myLanguage)
    {
      switch (myLanguage)
      {
        case "English":
          frenchToolStripMenuItem.Checked = false;
          englishToolStripMenuItem.Checked = true;
          fileToolStripMenuItem.Text = _languageDicoEn["MenuFile"];
          newToolStripMenuItem.Text = _languageDicoEn["MenuFileNew"];
          openToolStripMenuItem.Text = _languageDicoEn["MenuFileOpen"];
          saveToolStripMenuItem.Text = _languageDicoEn["MenuFileSave"];
          saveasToolStripMenuItem.Text = _languageDicoEn["MenuFileSaveAs"];
          printPreviewToolStripMenuItem.Text = _languageDicoEn["MenuFilePrint"];
          printPreviewToolStripMenuItem.Text = _languageDicoEn["MenufilePageSetup"];
          quitToolStripMenuItem.Text = _languageDicoEn["MenufileQuit"];
          editToolStripMenuItem.Text = _languageDicoEn["MenuEdit"];
          cancelToolStripMenuItem.Text = _languageDicoEn["MenuEditCancel"];
          redoToolStripMenuItem.Text = _languageDicoEn["MenuEditRedo"];
          cutToolStripMenuItem.Text = _languageDicoEn["MenuEditCut"];
          copyToolStripMenuItem.Text = _languageDicoEn["MenuEditCopy"];
          pasteToolStripMenuItem.Text = _languageDicoEn["MenuEditPaste"];
          selectAllToolStripMenuItem.Text = _languageDicoEn["MenuEditSelectAll"];
          toolsToolStripMenuItem.Text = _languageDicoEn["MenuTools"];
          personalizeToolStripMenuItem.Text = _languageDicoEn["MenuToolsCustomize"];
          optionsToolStripMenuItem.Text = _languageDicoEn["MenuToolsOptions"];
          languagetoolStripMenuItem.Text = _languageDicoEn["MenuLanguage"];
          englishToolStripMenuItem.Text = _languageDicoEn["MenuLanguageEnglish"];
          frenchToolStripMenuItem.Text = _languageDicoEn["MenuLanguageFrench"];
          helpToolStripMenuItem.Text = _languageDicoEn["MenuHelp"];
          summaryToolStripMenuItem.Text = _languageDicoEn["MenuHelpSummary"];
          indexToolStripMenuItem.Text = _languageDicoEn["MenuHelpIndex"];
          searchToolStripMenuItem.Text = _languageDicoEn["MenuHelpSearch"];
          aboutToolStripMenuItem.Text = _languageDicoEn["MenuHelpAbout"];
          DisplayToolStripMenuItem.Text = _languageDicoEn["Display"];
          SmallToolStripMenuItem.Text = _languageDicoEn["Small"];
          MediumToolStripMenuItem.Text = _languageDicoEn["Medium"];
          LargeToolStripMenuItem.Text = _languageDicoEn["Large"];


          _currentLanguage = "English";
          break;
        case "French":
          frenchToolStripMenuItem.Checked = true;
          englishToolStripMenuItem.Checked = false;
          fileToolStripMenuItem.Text = _languageDicoFr["MenuFile"];
          newToolStripMenuItem.Text = _languageDicoFr["MenuFileNew"];
          openToolStripMenuItem.Text = _languageDicoFr["MenuFileOpen"];
          saveToolStripMenuItem.Text = _languageDicoFr["MenuFileSave"];
          saveasToolStripMenuItem.Text = _languageDicoFr["MenuFileSaveAs"];
          printPreviewToolStripMenuItem.Text = _languageDicoFr["MenuFilePrint"];
          printPreviewToolStripMenuItem.Text = _languageDicoFr["MenufilePageSetup"];
          quitToolStripMenuItem.Text = _languageDicoFr["MenufileQuit"];
          editToolStripMenuItem.Text = _languageDicoFr["MenuEdit"];
          cancelToolStripMenuItem.Text = _languageDicoFr["MenuEditCancel"];
          redoToolStripMenuItem.Text = _languageDicoFr["MenuEditRedo"];
          cutToolStripMenuItem.Text = _languageDicoFr["MenuEditCut"];
          copyToolStripMenuItem.Text = _languageDicoFr["MenuEditCopy"];
          pasteToolStripMenuItem.Text = _languageDicoFr["MenuEditPaste"];
          selectAllToolStripMenuItem.Text = _languageDicoFr["MenuEditSelectAll"];
          toolsToolStripMenuItem.Text = _languageDicoFr["MenuTools"];
          personalizeToolStripMenuItem.Text = _languageDicoFr["MenuToolsCustomize"];
          optionsToolStripMenuItem.Text = _languageDicoFr["MenuToolsOptions"];
          languagetoolStripMenuItem.Text = _languageDicoFr["MenuLanguage"];
          englishToolStripMenuItem.Text = _languageDicoFr["MenuLanguageEnglish"];
          frenchToolStripMenuItem.Text = _languageDicoFr["MenuLanguageFrench"];
          helpToolStripMenuItem.Text = _languageDicoFr["MenuHelp"];
          summaryToolStripMenuItem.Text = _languageDicoFr["MenuHelpSummary"];
          indexToolStripMenuItem.Text = _languageDicoFr["MenuHelpIndex"];
          searchToolStripMenuItem.Text = _languageDicoFr["MenuHelpSearch"];
          aboutToolStripMenuItem.Text = _languageDicoFr["MenuHelpAbout"];
          DisplayToolStripMenuItem.Text = _languageDicoFr["Display"];
          SmallToolStripMenuItem.Text = _languageDicoFr["Small"];
          MediumToolStripMenuItem.Text = _languageDicoFr["Medium"];
          LargeToolStripMenuItem.Text = _languageDicoFr["Large"];

          _currentLanguage = "French";
          break;
        default:
          SetLanguage("English");
          break;
      }
    }

    private void cutToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Control focusedControl = FindFocusedControl(new List<Control>
      {
        textBoxCsvFilePath, textBoxDecryptFilePath, textBoxWebBrowserUrl
      });

      var tb = focusedControl as TextBox;
      if (tb != null)
      {
        CutToClipboard(tb);
      }
    }

    private void copyToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Control focusedControl = FindFocusedControl(new List<Control>
      {
        textBoxCsvFilePath, textBoxDecryptFilePath, textBoxWebBrowserUrl
      });

      var tb = focusedControl as TextBox;
      if (tb != null)
      {
        CopyToClipboard(tb);
      }
    }

    private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Control focusedControl = FindFocusedControl(new List<Control>
      {
        textBoxCsvFilePath, textBoxDecryptFilePath, textBoxWebBrowserUrl
      });

      var tb = focusedControl as TextBox;
      if (tb != null)
      {
        PasteFromClipboard(tb);
      }
    }

    private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Control focusedControl = FindFocusedControl(new List<Control>
      {
        textBoxCsvFilePath, textBoxDecryptFilePath, textBoxWebBrowserUrl
      });

      TextBox control = focusedControl as TextBox;
      if (control != null) control.SelectAll();
    }

    private void CutToClipboard(TextBoxBase tb, string errorMessage = "nothing")
    {
      if (tb != ActiveControl) return;
      if (tb.Text == string.Empty)
      {
        DisplayMessage(Translate("ThereIs") + Punctuation.OneSpace +
          Translate(errorMessage) + Punctuation.OneSpace +
          Translate("ToCut") + Punctuation.OneSpace, Translate(errorMessage),
          MessageBoxButtons.OK);
        return;
      }

      if (tb.SelectedText == string.Empty)
      {
        DisplayMessage(Translate("NoTextHasBeenSelected"),
          Translate(errorMessage), MessageBoxButtons.OK);
        return;
      }

      Clipboard.SetText(tb.SelectedText);
      tb.SelectedText = string.Empty;
    }

    private void CopyToClipboard(TextBoxBase tb, string message = "nothing")
    {
      if (tb != ActiveControl) return;
      if (tb.Text == string.Empty)
      {
        DisplayMessage(Translate("ThereIsNothingToCopy") + Punctuation.OneSpace,
          Translate(message), MessageBoxButtons.OK);
        return;
      }

      if (tb.SelectedText == string.Empty)
      {
        DisplayMessage(Translate("NoTextHasBeenSelected"),
          Translate(message), MessageBoxButtons.OK);
        return;
      }

      Clipboard.SetText(tb.SelectedText);
    }

    private void PasteFromClipboard(TextBoxBase tb)
    {
      if (tb != ActiveControl) return;
      var selectionIndex = tb.SelectionStart;
      tb.SelectedText = Clipboard.GetText();
      tb.SelectionStart = selectionIndex + Clipboard.GetText().Length;
    }

    private void DisplayMessage(string message, string title, MessageBoxButtons buttons)
    {
      MessageBox.Show(this, message, title, buttons);
    }

    private string Translate(string index)
    {
      string result = string.Empty;
      switch (_currentLanguage.ToLower())
      {
        case "english":
          result = _languageDicoEn.ContainsKey(index) ? _languageDicoEn[index] :
           "the term: \"" + index + "\" has not been translated yet.\nPlease translate this term";
          break;
        case "french":
          result = _languageDicoFr.ContainsKey(index) ? _languageDicoFr[index] :
            "the term: \"" + index + "\" has not been translated yet.\nPlease translate this term";
          break;
      }

      return result;
    }

    private static Control FindFocusedControl(Control container)
    {
      foreach (Control childControl in container.Controls.Cast<Control>().Where(childControl => childControl.Focused))
      {
        return childControl;
      }

      return (from Control childControl in container.Controls
              select FindFocusedControl(childControl)).FirstOrDefault(maybeFocusedControl => maybeFocusedControl != null);
    }

    private static Control FindFocusedControl(List<Control> container)
    {
      return container.FirstOrDefault(control => control.Focused);
    }

    private static Control FindFocusedControl(params Control[] container)
    {
      return container.FirstOrDefault(control => control.Focused);
    }

    private static Control FindFocusedControl(IEnumerable<Control> container)
    {
      return container.FirstOrDefault(control => control.Focused);
    }

    private static string PeekDirectory()
    {
      string result = string.Empty;
      FolderBrowserDialog fbd = new FolderBrowserDialog();
      if (fbd.ShowDialog() == DialogResult.OK)
      {
        result = fbd.SelectedPath;
      }

      return result;
    }

    private string PeekFile()
    {
      string result = string.Empty;
      OpenFileDialog fd = new OpenFileDialog();
      if (fd.ShowDialog() == DialogResult.OK)
      {
        result = fd.SafeFileName;
      }

      return result;
    }

    private static string PeekFile(string initialDirectory, string filters)
    {
      string result = string.Empty;
      OpenFileDialog fileDialog = new OpenFileDialog();
      if (initialDirectory == string.Empty)
      {
        initialDirectory = "C:\\";
      }

      if (filters == string.Empty)
      {
        filters = Settings.Default.FilterOpenFiles;
      }

      //openFileDialog1.InitialDirectory = initialDirectory;
      fileDialog.Filter = filters;
      //openFileDialog1.FilterIndex = 2;
      fileDialog.RestoreDirectory = true;

      if (fileDialog.ShowDialog() == DialogResult.OK)
      {
        result = fileDialog.FileName;
      }

      return result;
    }

    private void SmallToolStripMenuItem_Click(object sender, EventArgs e)
    {
      UncheckAllOptions();
      SmallToolStripMenuItem.Checked = true;
      AdjustAllControls();
    }

    private void MediumToolStripMenuItem_Click(object sender, EventArgs e)
    {
      UncheckAllOptions();
      MediumToolStripMenuItem.Checked = true;
      AdjustAllControls();
    }

    private void LargeToolStripMenuItem_Click(object sender, EventArgs e)
    {
      UncheckAllOptions();
      LargeToolStripMenuItem.Checked = true;
      AdjustAllControls();
    }

    private static void AdjustControls(params Control[] listOfControls)
    {
      if (listOfControls.Length == 0)
      {
        return;
      }

      int position = listOfControls[0].Width + 33; // 33 is the initial padding
      bool isFirstControl = true;
      foreach (Control control in listOfControls)
      {
        if (isFirstControl)
        {
          isFirstControl = false;
        }
        else
        {
          control.Left = position + 10;
          position += control.Width;
        }
      }
    }

    private void AdjustAllControls()
    {
      AdjustControls(); // insert here all labels, textboxes and buttons, one method per line of controls
    }

    private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      FormOptions frmOptions = new FormOptions(_configurationOptions);

      if (frmOptions.ShowDialog() == DialogResult.OK)
      {
        _configurationOptions = frmOptions.ConfigurationOptions2;
      }
    }

    private static void SetButtonEnabled(Button button, params Control[] controls)
    {
      bool result = true;
      foreach (Control ctrl in controls)
      {
        if (ctrl.GetType() == typeof(TextBox))
        {
          if (((TextBox)ctrl).Text == string.Empty)
          {
            result = false;
            break;
          }
        }

        if (ctrl.GetType() == typeof(ListView))
        {
          if (((ListView)ctrl).Items.Count == 0)
          {
            result = false;
            break;
          }
        }

        if (ctrl.GetType() == typeof(ComboBox))
        {
          if (((ComboBox)ctrl).SelectedIndex == -1)
          {
            result = false;
            break;
          }
        }
      }

      button.Enabled = result;
    }

    private void textBoxName_KeyDown(object sender, KeyEventArgs e)
    {
      if (e.KeyCode == Keys.Enter)
      {
        // do something
      }
    }

    private void buttonCsvFilePathPeek_Click(object sender, EventArgs e)
    {
      // Choisir un fichier csv uniquement.
      string fileName = PeekFile(string.Empty, "CSV files (*.csv)|*.csv");
      if (fileName != string.Empty)
      {
        textBoxCsvFilePath.Text = fileName;
      }


    }

    private void buttonDecryptFilePathPeek_Click(object sender, EventArgs e)
    {
      // Choisir l'emplacement du programme Decrypt/Encrypt.
      string fileName = PeekFile(string.Empty, "EXE files (*.exe)|*.exe");
      if (fileName != string.Empty)
      {
        textBoxDecryptFilePath.Text = fileName;
      }
    }

    private void buttonStartDecryptProgram_Click(object sender, EventArgs e)
    {
      if (File.Exists(textBoxDecryptFilePath.Text))
      {
        Process.Start(textBoxDecryptFilePath.Text);
      }
    }

    private void buttonCheclAll_Click(object sender, EventArgs e)
    {
      if (dataGridViewMain.ColumnCount != 0)
      {
        CheckAllItems(dataGridViewMain);
      }
    }

    private static void CheckAllItems(DataGridView dgw)
    {
      dgw.Rows.OfType<DataGridViewRow>().ToList().ForEach(row => row.Cells[0].Value = true);
    }

    private static void UnCheckAllItems(DataGridView dgw)
    {
      dgw.Rows.OfType<DataGridViewRow>().ToList().ForEach(row => row.Cells[0].Value = false);
    }

    private static void ToggleAllItems(DataGridView dgw)
    {
      dgw.Rows.OfType<DataGridViewRow>().ToList().ForEach(row => row.Cells[0].Value = !(bool)row.Cells[0].Value);
    }

    private void buttonUncheckAll_Click(object sender, EventArgs e)
    {
      if (dataGridViewMain.ColumnCount != 0)
      {
        UnCheckAllItems(dataGridViewMain);
      }
    }

    private void buttonToggleAll_Click(object sender, EventArgs e)
    {
      if (dataGridViewMain.ColumnCount != 0)
      {
        ToggleAllItems(dataGridViewMain);
      }
    }

    private void buttonDataGridViewLoad_Click(object sender, EventArgs e)
    {
      if (!File.Exists(textBoxCsvFilePath.Text)) return;
      // Lecture du fichier csv
      dataTableCsvFile = new DataTable();
      dataTableCsvFile = CreateDataTableCsvFile();
      FeedDataTableCsvFile(dataTableCsvFile);
      dataGridViewMain.DataSource = dataTableCsvFile;
      dataGridViewMain.Columns[0].Width = Settings.Default.DgvSizeColum0;
      dataGridViewMain.Columns[1].Width = Settings.Default.DgvSizeColum1;
      dataGridViewMain.Columns[2].Width = Settings.Default.DgvSizeColum2;
      dataGridViewMain.Columns[3].Width = Settings.Default.DgvSizeColum3;
      dataGridViewMain.Columns[4].Width = Settings.Default.DgvSizeColum4;
      dataGridViewMain.Columns[5].Width = Settings.Default.DgvSizeColum5;
      dataGridViewMain.Columns[6].Width = Settings.Default.DgvSizeColum6;
      dataGridViewMain.Columns[7].Width = Settings.Default.DgvSizeColum7;
      dataGridViewMain.Columns[8].Width = Settings.Default.DgvSizeColum8;
      dataGridViewMain.Columns[9].Width = Settings.Default.DgvSizeColum9;
      dataGridViewMain.Columns[10].Width = Settings.Default.DgvSizeColum10;
      dataGridViewMain.Columns[11].Width = Settings.Default.DgvSizeColum11;
      PrepareCells();
    }

    private void FeedDataTableCsvFile(DataTable dt)
    {
      // on rempli le datatable dt
      // lecture du fichier csv

      foreach (var line in ReadCsvFile(textBoxCsvFilePath.Text))
      {
        DataRow row = dt.NewRow();
        // Vérifier
        row[0] = false;
        // Environnement
        row[1] = line[0];
        // URL
        row[2] = line[2];
        // Ping
        row[3] = string.Empty;
        // HTTP
        row[4] = string.Empty;
        // HTML
        row[5] = string.Empty;
        // Login
        row[6] = string.Empty;
        // Web
        row[7] = string.Empty;
        // Erreur
        row[8] = string.Empty;
        // Server
        row[9] = line[1];
        // UserName
        row[10] = line[3];
        // Password
        row[11] = line[4];
        // DSI
        row[12] = line[5];
        // Application
        row[13] = line[6];
        // Responsable projet
        row[14] = line[7];

        dt.Rows.Add(row);
      }
    }

    private static IEnumerable<string[]> ReadCsvFile(string fileName)
    {

      if (!StringHelpers.IniFileExists(fileName))
      {
        try
        {
          CreateIniFile(fileName);
        }
        catch (Exception exception)
        {
          MessageBox.Show(Resources.Une_erreur_s_est_produite_lors_de_la_création_du_fichier_ini + Punctuation.OneSpace + exception.Message);
          return new List<string[]>();
        }
      }

      List<string[]> arrayList = new List<string[]>();
      string connexionString = string.Format(CultureInfo.CurrentCulture,
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Mode=Read;Extended Properties=\"TEXT;HDR=YES;FMT=Delimited\"",
        StringHelpers.AddBackSlash(StringHelpers.GetDirectoryName(fileName))
        );

      try
      {
        using (OleDbConnection dbConnection = new OleDbConnection(connexionString))
        {
          dbConnection.Open();
          string requete_import = "Select * from [" + StringHelpers.ExtractFileName(fileName) + "]";
          OleDbCommand sourceCommande = new OleDbCommand(requete_import, dbConnection);
          using (DbDataReader sourceReader = sourceCommande.ExecuteReader())
          {
            Debug.Assert(sourceReader != null, "sourceReader != null");
            while (sourceReader.Read())
            {
              arrayList.Add(new[]
              {
                // environnement
                sourceReader.GetValue(0).ToString(), 
                // serveur
                sourceReader.GetValue(1).ToString(),
                // URL
                sourceReader.GetValue(4).ToString(), 
                // User name
                sourceReader.GetValue(10).ToString(),
                // Password
                sourceReader.GetValue(11).ToString(),
                // DSI
                sourceReader.GetValue(25).ToString(), 
                // Application
                sourceReader.GetValue(26).ToString(),
                // Responsable projet
                sourceReader.GetValue(28).ToString()
              });
            }
          }
        }
      }
      catch (Exception exception)
      {
        MessageBox.Show(Resources.Une_erreur_s_est_produite_lors_de_la_lecture_du_fichier_csv + Punctuation.OneSpace + exception.Message);
        return new List<string[]>();
      }

      return arrayList;
    }

    private static void CreateIniFile(string fileName)
    {
      try
      {
        string DirectoryFileName = Path.GetDirectoryName(fileName);
        if (!File.Exists(Settings.Default.SchemaIniFileName))
        {
          if (DirectoryFileName != null)
            using (StreamWriter sw = new StreamWriter(Path.Combine(DirectoryFileName, Settings.Default.SchemaIniFileName)))
            {
              sw.WriteLine("[" + Path.GetFileName(fileName) + "]");
              sw.WriteLine("ColNameHeader=True");
              sw.WriteLine("CharacterSet=ANSI");
              sw.WriteLine("Format=Delimited(;)");
              sw.WriteLine("MaxScanRows=0");
            }
        }
      }
      catch (Exception exception)
      {
        MessageBox.Show(
          Resources.Impossible_de_créer_le_fichier + Punctuation.OneSpace + Settings.Default.SchemaIniFileName +
          Punctuation.OneSpace + Resources.dans_le_répertoire + Punctuation.OneSpace + Path.GetDirectoryName(fileName) +
          Punctuation.OneSpace + exception.Message);
      }
    }

    private static DataTable CreateDataTableCsvFile()
    {
      // Create new DataTable.
      DataTable table = new DataTable();

      // Declare DataColumn and DataRow variables.
      DataColumn column;

      // Create new DataColumn, set DataType, ColumnName
      // and add to DataTable.    
      column = new DataColumn();
      column.DataType = Type.GetType("System.Boolean");
      column.ColumnName = "Vérifier";
      table.Columns.Add(column);

      // Create second column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Environnement";
      table.Columns.Add(column);

      // Create URL column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "URL";
      table.Columns.Add(column);

      // Create Ping column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Ping";
      table.Columns.Add(column);

      // Create HTTP column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "HTTP";
      table.Columns.Add(column);

      // Create HTML column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "HTML";
      table.Columns.Add(column);

      // Create Login column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Login";
      table.Columns.Add(column);

      // Create WEB column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "WEB";
      table.Columns.Add(column);

      // Create Erreur column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Erreur";
      table.Columns.Add(column);

      // Create Server column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Server";
      table.Columns.Add(column);

      // Create UserName column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "UserName";
      table.Columns.Add(column);

      // Create Password column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Password";
      table.Columns.Add(column);

      // Create DSI column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "DSI";
      table.Columns.Add(column);

      // Create Application column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Application";
      table.Columns.Add(column);

      // Create Responsable column.
      column = new DataColumn();
      column.DataType = Type.GetType("System.String");
      column.ColumnName = "Responsable";
      table.Columns.Add(column);

      return table;
    }

    private void AdjustDataGridViewSizingToAuto()
    {
      dataGridViewMain.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
    }

    private void buttonDataGridviewVerify_Click(object sender, EventArgs e)
    {
      // on vérifie chaque dispatcher qui a été coché
      // on vérifie le login et si ok alors ping ok et http ok
      if (dataGridViewMain.Rows.Count == 0)
      {
        MessageBox.Show(Resources.Vous_devez_lancer_le_chargement_du_fichier_avant_la_vérification);
        return;
      }

      labelPleaseWait.Visible = true;
      progressBarData.Visible = true;
      buttonDataGridviewVerify.Enabled = false;
      Application.DoEvents();

      // traitement de login d'abord puis http puis ping
      progressBarData.Minimum = 1;
      progressBarData.Value = progressBarData.Minimum;
      progressBarData.Maximum = dataTableCsvFile.Rows.Count + 1;
      int counter = progressBarData.Minimum;
      //ServicePointManager.DefaultConnectionLimit = 2;
      foreach (DataRow row in dataTableCsvFile.Rows)
      {
        // on traite seulement les lignes cochées
        if (row[0].ToString() == true.ToString())
        {
          progressBarData.Value = ++counter;

          //Uri baseUri = new Uri(row["URL"].ToString());
          //Uri htmlUri = new Uri(baseUri, "/ext?encoding=UTF-8&b_action=p2plbDiag");

          string url = row["URL"].ToString();
          if (url.EndsWith("/"))
          {
            url += "ext?encoding=UTF-8&b_action=p2plbDiag";
          }
          else
          {
            url += "/ext?encoding=UTF-8&b_action=p2plbDiag";
          }

          string userName = row["UserName"].ToString();
          string password = Functions.Decrypt(row["Password"].ToString());
          string error;
          string httpResult = GetHttpResponse(url, out error);
          row["Erreur"] = string.Empty;

          if (httpResult.Contains("ID utilisateur") && httpResult.Contains("Mot de passe"))
          {
            row["HTTP"] = "OK";
            row["HTML"] = "OK";
            row["Ping"] = "OK";
            Application.DoEvents();
            httpResult = TryHttpRequestWithLogon(url, userName, password);
            if (httpResult.StartsWith("erreur:"))
            {
              row["Erreur"] = httpResult;
            }

            if (httpResult.Contains("logoff") && httpResult.Contains("Se déconnecter"))
            {
              row["Login"] = "OK";
              row["HTTP"] = "OK";
              row["HTML"] = "OK";
              row["Ping"] = "OK";
              Application.DoEvents();
            }
            else
            {
              row["Login"] = "KO";
              Application.DoEvents();
            }
          }
          else
          {
            if (httpResult != string.Empty)
            {
              row["HTTP"] = "OK";
              row["HTML"] = "KO";
              row["Ping"] = "OK";
              row["Login"] = "KO";
              Application.DoEvents();
            }
            else
            {
              row["HTTP"] = "KO";
              row["HTML"] = "KO";
              row["Login"] = "KO";
              row["Erreur"] += error + "\n";
              Application.DoEvents();

              // sinon on essaye le ping
              string urlServerName = StringHelpers.GetServerNameFromUrl(row[2].ToString());

              Ping pingSender = new Ping();

              try
              {
                PingReply reply = pingSender.Send(urlServerName, 120);
                if (reply.Status == IPStatus.Success)
                {
                  row["Ping"] = "OK";
                }
                else
                {
                  row["Ping"] = "KO";
                }
              }
              catch (Exception ex)
              {
                row["Erreur"] += ex.Message + "\n";
                row["Ping"] = "KO";
              }

              Application.DoEvents();
            }
          }

          Application.DoEvents();
        }
      }

      ColorCells();
      AddButtons();
      labelPleaseWait.Visible = false;
      progressBarData.Value = progressBarData.Minimum;
      progressBarData.Visible = false;
      buttonDataGridviewVerify.Enabled = true;
    }

    private void AddButtons()
    {
      // ajout d'un bouton ou url pour copier l'url ko vers l'onglet Web
      foreach (DataGridViewRow row in dataGridViewMain.Rows)
      {
        if (row.Cells["Login"].Value.ToString().ToUpper() == "KO")
        {
          DataGridViewButtonCell button = new DataGridViewButtonCell();
          button.ToolTipText = "WEB";
          row.Cells["WEB"] = button;
          row.Cells["WEB"].Value = "Web";
        }
      }
    }

    private void ColorCells()
    {
      DataGridViewCellStyle redCellStyle = new DataGridViewCellStyle();
      redCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
      redCellStyle.Font = new Font(dataGridViewMain.Font, FontStyle.Bold);
      redCellStyle.ForeColor = Color.Red;

      DataGridViewCellStyle greenCellStyle = new DataGridViewCellStyle();
      greenCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
      greenCellStyle.Font = new Font(dataGridViewMain.Font, FontStyle.Bold);
      greenCellStyle.ForeColor = Color.Green;

      foreach (DataGridViewRow row in dataGridViewMain.Rows)
      {
        if (row.Cells["Ping"].Value.ToString().ToUpper() == "OK")
        {
          row.Cells["Ping"].Style = greenCellStyle;
        }
        else if (row.Cells["Ping"].Value.ToString().ToUpper() == "KO")
        {
          row.Cells["Ping"].Style = redCellStyle;
        }

        if (row.Cells["HTTP"].Value.ToString().ToUpper() == "OK")
        {
          row.Cells["HTTP"].Style = greenCellStyle;
        }
        else if (row.Cells["HTTP"].Value.ToString().ToUpper() == "KO")
        {
          row.Cells["HTTP"].Style = redCellStyle;
        }

        if (row.Cells["HTML"].Value.ToString().ToUpper() == "OK")
        {
          row.Cells["HTML"].Style = greenCellStyle;
        }
        else if (row.Cells["HTML"].Value.ToString().ToUpper() == "KO")
        {
          row.Cells["HTML"].Style = redCellStyle;
        }

        if (row.Cells["Login"].Value.ToString().ToUpper() == "OK")
        {
          row.Cells["Login"].Style = greenCellStyle;
        }
        else if (row.Cells["Login"].Value.ToString().ToUpper() == "KO")
        {
          row.Cells["Login"].Style = redCellStyle;
        }
      }
    }

    private void PrepareCells()
    {
      DataGridViewCellStyle centerBoldCellStyle = new DataGridViewCellStyle();
      centerBoldCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
      centerBoldCellStyle.Font = new Font(dataGridViewMain.Font, FontStyle.Bold);

      foreach (DataGridViewRow row in dataGridViewMain.Rows)
      {
        row.Cells[3].Style = centerBoldCellStyle;
        row.Cells[4].Style = centerBoldCellStyle;
        row.Cells[5].Style = centerBoldCellStyle;
        row.Cells[6].Style = centerBoldCellStyle;
        row.Cells[7].Style = centerBoldCellStyle;
      }
    }

    private void dataGridViewMain_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
      if (e.ColumnIndex == 0)
      {
        int numberOfRow = dataTableCsvFile.AsEnumerable().Count(r => r[0].ToString() == true.ToString());
        buttonDataGridviewVerify.Enabled = numberOfRow > 0;
      }
    }

    private async void getResult(string url)
    {
      HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
      HttpWebResponse response = (HttpWebResponse)await request.GetResponseAsync();
      string resultString;
      //get the string from the response
      using (var xx = response.GetResponseStream())
      {
        StreamReader reader = new StreamReader(xx, Encoding.UTF8);
        resultString = await reader.ReadToEndAsync();
      }

      //return new Task();
    }

    private static string TryHttpRequestWithLogon(string url, string userName, string passsword)
    {
      string result = string.Empty;
      // TODO modifier l'url pour faire une requête avec login et password
      string uriStr = StringHelpers.AddSlash(url) + "?username=" + HttpUtility.UrlEncode(userName) + "&password=" + HttpUtility.UrlEncode(passsword);
      Uri address = new Uri(uriStr);
      HttpWebRequest request = WebRequest.Create(address) as HttpWebRequest;

      // Get response 
      if (request != null)
      {
        try
        {
          using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
          {
            // Get the response stream  
            if (response != null)
            {
              using (StreamReader reader = new StreamReader(response.GetResponseStream()))
              {
                result = reader.ReadToEnd();
              }
            }
          }
        }
        catch (Exception exception)
        {
          // ignored
          result = "error: " + exception.Message;
        }
      }

      return result;
    }

    private static string TryHttpRequest(string url)
    {
      string result = string.Empty;
      Uri address = new Uri(url);
      HttpWebRequest request = WebRequest.Create(address) as HttpWebRequest;
      request.Timeout = 1000;
      request.Proxy = null;

      // Get response 
      if (request != null)
      {
        try
        {
          using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
          {
            // Get the response stream  
            if (response != null)
            {
              using (StreamReader reader = new StreamReader(response.GetResponseStream()))
              {
                result = reader.ReadToEnd();
              }
            }
          }
        }
        catch (Exception exception)
        {
          // ignored
          //MessageBox.Show(exception.Message);
          request.Abort();
          request = null;
        }
      }

      return result;
    }

    public string GetHttpResponse(string URL, out string error)
    {
      error = string.Empty;

      HttpWebRequest request = WebRequest.Create(URL) as HttpWebRequest;
      ////request.Credentials = new NetworkCredential(Settings.Default.LicenseUser, Settings.Default.LicensePassword);
      request.KeepAlive = false;
      request.Timeout = 5000;
      request.Proxy = null;

      request.ServicePoint.ConnectionLeaseTimeout = 5000;
      request.ServicePoint.MaxIdleTime = 5000;

      // Read stream
      string responseString = string.Empty;
      try
      {
        using (WebResponse response = request.GetResponse())
        {
          using (Stream objStream = response.GetResponseStream())
          {
            using (StreamReader objReader = new StreamReader(objStream))
            {
              responseString = objReader.ReadToEnd();
              objReader.Close();
            }
            objStream.Flush();
            objStream.Close();
          }
          response.Close();
        }
      }
      catch (WebException exception)
      {
        // ignored
        //MessageBox.Show("erreur" + exception.Message);
        error = exception.Message;
      }
      finally
      {
        request.Abort();
      }

      return responseString;
    }


    private void dataGridViewMain_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
    {
      if (e.Button == MouseButtons.Right)
      {
        contextMenuStripDgv.Show(Cursor.Position);
      }
    }

    private WebRequest CreateSoapRequest(ISoapMessage soapMessage)
    {
      var wr = WebRequest.Create(soapMessage.Uri);
      wr.ContentType = "text/xml;charset=utf-8";
      wr.ContentLength = soapMessage.ContentXml.Length;
      wr.Headers.Add("SOAPAction", soapMessage.SoapAction);
      wr.Credentials = soapMessage.Credentials;
      wr.Method = "POST";
      wr.GetRequestStream().Write(Encoding.UTF8.GetBytes(soapMessage.ContentXml), 0, soapMessage.ContentXml.Length);

      return wr;
    }

    public interface ISoapMessage
    {
      string Uri { get; }
      string ContentXml { get; }
      string SoapAction { get; }
      ICredentials Credentials { get; }
    }

    private void buttonWebBrowserGo_Click(object sender, EventArgs e)
    {
      if (textBoxWebBrowserUrl.Text == string.Empty)
      {
        MessageBox.Show(Resources.L_url_est_manquante);
        return;
      }

      webBrowserMain.Navigate(textBoxWebBrowserUrl.Text);
    }

    private void webBrowserMain_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
    {
      // when the web browser document is loaded completely
      // search for CAMUsername for username
      // and  for password
      // and cmdOK for button OK

    }

    private void buttonWebBrowserGoLogin_Click(object sender, EventArgs e)
    {
      if (webBrowserMain.Document != null)
      {
        webBrowserMain.ScriptErrorsSuppressed = true;
        HtmlElement userNameElementById = webBrowserMain.Document.GetElementById("Username");
        HtmlElement PasswordElementById = webBrowserMain.Document.GetElementById("Password");
        HtmlElement OkButtonElementById = webBrowserMain.Document.GetElementById("cmdOK");

        if (userNameElementById == null || PasswordElementById == null || OkButtonElementById == null)
        {
          return;
        }

        userNameElementById.SetAttribute("Value", GetUserName(textBoxWebBrowserUrl.Text));
        PasswordElementById.SetAttribute("Value", GetPassword(textBoxWebBrowserUrl.Text));
        OkButtonElementById.InvokeMember("click");
      }
    }

    private string GetUserName(string text)
    {
      string result = string.Empty;

      return result;
    }

    private string GetPassword(string text)
    {
      string result = string.Empty;

      return result;
    }

    private void contextMenuStripDgv_MouseClick(object sender, MouseEventArgs e)
    {
      if (dataGridViewMain.SelectedCells.Count == 1)
      {
        // une seule sélection
        if (dataGridViewMain.CurrentCell.OwningColumn.Name == "Password")
        {
          Clipboard.SetText(Functions.Decrypt(dataGridViewMain.CurrentCell.Value.ToString()));
        }
        else
        {
          Clipboard.SetText(dataGridViewMain.CurrentCell.Value.ToString());
        }
      }
      else
      {
        // plusieures cellules sélectionnées
        string allCellsValue = string.Empty;
        for (int i = 0; i < dataGridViewMain.SelectedCells.Count; i++)
        {
          allCellsValue += dataGridViewMain.SelectedCells[i].Value + Environment.NewLine;
        }

        Clipboard.SetText(allCellsValue);
      }

    }

    private void buttonSendToWebBrowser_Click(object sender, EventArgs e)
    {
      if (dataGridViewMain.SelectedCells.Count == 1)
      {
        string cellValue = dataGridViewMain.CurrentCell.Value.ToString();
        int currentRow = 1;
        tabControlMain.SelectedIndex = 2;
        textBoxWebBrowserUrl.Text = string.Empty;
      }
    }

    private void dataGridViewMain_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {
      var senderGrid = (DataGridView)sender;
      DataGridViewColumn dataGridViewColumn = dataGridViewMain.Columns["Web"];
      if (dataGridViewColumn == null || e.ColumnIndex != dataGridViewColumn.Index || e.RowIndex < 0)
      {
        return;
      }

      textBoxWebBrowserUrl.Text = dataGridViewMain.Rows[e.RowIndex].Cells["URL"].Value.ToString();
      tabControlMain.SelectedIndex = 2;
    }
  }
}