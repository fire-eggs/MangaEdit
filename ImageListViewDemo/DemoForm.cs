using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Net;
using SevenZip;
using System.Drawing.Imaging;
using System.Linq;

namespace Manina.Windows.Forms
{
    public partial class DemoForm : Form
    {
        #region Member variables
        private BackgroundWorker bw = new BackgroundWorker();
        #endregion

        #region Renderer and color combobox items
        /// <summary>
        /// Represents an item in the renderer combobox.
        /// </summary>
        private struct RendererComboBoxItem
        {
            public string Name;
            public string FullName;

            public override string ToString()
            {
                return Name;
            }

            public RendererComboBoxItem(Type type)
            {
                Name = type.Name;
                FullName = type.FullName;
            }
        }

        /// <summary>
        /// Represents an item in the custom color combobox.
        /// </summary>
        private struct ColorComboBoxItem
        {
            public string Name;
            public PropertyInfo Field;

            public override string ToString()
            {
                return Name;
            }

            public ColorComboBoxItem(PropertyInfo field)
            {
                Name = field.Name;
                Field = field;
            }
        }
        #endregion

        #region Constructor
        public DemoForm()
        {
            InitializeComponent();

            paneToolStripButton_Click(null, null); // start in pane mode
            LoadSettings();

            // Setup the background worker
            Application.Idle += new EventHandler(Application_Idle);
            //bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            //bw.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            // Find and add custom colors
            Type colorType = typeof(ImageListViewColor);
            int i = 0;
            foreach (PropertyInfo field in colorType.GetProperties(BindingFlags.Public | BindingFlags.Static))
            {
                colorToolStripComboBox.Items.Add(new ColorComboBoxItem(field));
                if (field.Name == "Default")
                    colorToolStripComboBox.SelectedIndex = i;
                i++;
            }
            // Dynamically add aligment values
            foreach (object o in Enum.GetValues(typeof(ContentAlignment)))
            {
                ToolStripMenuItem item1 = new ToolStripMenuItem(o.ToString());
                item1.Tag = o;
                item1.Click += new EventHandler(checkboxAlignmentToolStripButton_Click);
                checkboxAlignmentToolStripMenuItem.DropDownItems.Add(item1);
                ToolStripMenuItem item2 = new ToolStripMenuItem(o.ToString());
                item2.Tag = o;
                item2.Click += new EventHandler(iconAlignmentToolStripButton_Click);
                iconAlignmentToolStripMenuItem.DropDownItems.Add(item2);
            }

            imageListView1.AllowDuplicateFileNames = true;
            imageListView1.SetRenderer(new ImageListViewRenderers.DefaultRenderer());
            imageListView1.SortColumn = 0;
            imageListView1.SortOrder = SortOrder.AscendingNatural;
            string cacheDir = Path.Combine(Path.GetTempPath(), "Cache");
            if (!Directory.Exists(cacheDir))
                Directory.CreateDirectory(cacheDir);
            imageListView1.PersistentCacheDirectory = cacheDir;

            //imageListView1.Columns.Add(ColumnType.Name);
            //imageListView1.Columns.Add(ColumnType.Dimensions);
            //imageListView1.Columns.Add(ColumnType.FileSize);
            //imageListView1.Columns.Add(ColumnType.FolderName);

            TreeNode node = new TreeNode("Loading...", 3, 3);
            node.Tag = null;
            while (bw.IsBusy) ;
            bw.RunWorkerAsync(node);
        }

        #endregion

        #region Update UI while idle
        void Application_Idle(object sender, EventArgs e)
        {
            detailsToolStripButton.Checked = (imageListView1.View == View.Details);
            thumbnailsToolStripButton.Checked = (imageListView1.View == View.Thumbnails);
            galleryToolStripButton.Checked = (imageListView1.View == View.Gallery);
            paneToolStripButton.Checked = (imageListView1.View == View.Pane);

            integralScrollToolStripMenuItem.Checked = imageListView1.IntegralScroll;

            showCheckboxesToolStripMenuItem.Checked = imageListView1.ShowCheckBoxes;
            showFileIconsToolStripMenuItem.Checked = imageListView1.ShowFileIcons;

            x96ToolStripMenuItem.Checked = imageListView1.ThumbnailSize == new Size(96, 96);
            x120ToolStripMenuItem.Checked = imageListView1.ThumbnailSize == new Size(120, 120);
            x200ToolStripMenuItem.Checked = imageListView1.ThumbnailSize == new Size(200, 200);

            allowCheckBoxClickToolStripMenuItem.Checked = imageListView1.AllowCheckBoxClick;
            allowColumnClickToolStripMenuItem.Checked = imageListView1.AllowColumnClick;
            allowColumnResizeToolStripMenuItem.Checked = imageListView1.AllowColumnResize;
            allowPaneResizeToolStripMenuItem.Checked = imageListView1.AllowPaneResize;
            multiSelectToolStripMenuItem.Checked = imageListView1.MultiSelect;
            allowDragToolStripMenuItem.Checked = imageListView1.AllowDrag;
            allowDropToolStripMenuItem.Checked = imageListView1.AllowDrop;
            allowDuplicateFilenamesToolStripMenuItem.Checked = imageListView1.AllowDuplicateFileNames;
            continuousCacheModeToolStripMenuItem.Checked = (imageListView1.CacheMode == CacheMode.Continuous);

            ContentAlignment ca = imageListView1.CheckBoxAlignment;
            foreach (ToolStripMenuItem item in checkboxAlignmentToolStripMenuItem.DropDownItems)
                item.Checked = (ContentAlignment)item.Tag == ca;
            ContentAlignment ia = imageListView1.IconAlignment;
            foreach (ToolStripMenuItem item in iconAlignmentToolStripMenuItem.DropDownItems)
                item.Checked = (ContentAlignment)item.Tag == ia;

            string selname = "";
            if (imageListView1.SelectedItems.Count > 0)
            {
                selname = " (" + imageListView1.SelectedItems[0].Text + ")";
                Size sz = imageListView1.SelectedItems[0].Dimensions;
                selname += "[" + sz.Width + "w x" + sz.Height + "h]";
            }
            toolStripStatusLabel1.Text = string.Format("{0} Items: {1} Selected{3}, {2} Checked",
                imageListView1.Items.Count, imageListView1.SelectedItems.Count, imageListView1.CheckedItems.Count,
                selname);

            Text = "MangaEdit :" + Path.GetFileName(_zippath) + selname;

            groupAscendingToolStripMenuItem.Checked = imageListView1.GroupOrder == SortOrder.Ascending;
            groupDescendingToolStripMenuItem.Checked = imageListView1.GroupOrder == SortOrder.Descending;
            sortAscendingToolStripMenuItem.Checked = imageListView1.SortOrder == SortOrder.Ascending;
            sortDescendingToolStripMenuItem.Checked = imageListView1.SortOrder == SortOrder.Descending;
        }
        #endregion

        #region Set ImageListView options

        private void checkboxAlignmentToolStripButton_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            ContentAlignment aligment = (ContentAlignment)item.Tag;
            imageListView1.CheckBoxAlignment = aligment;
        }

        private void iconAlignmentToolStripButton_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            ContentAlignment aligment = (ContentAlignment)item.Tag;
            imageListView1.IconAlignment = aligment;
        }

        private void colorToolStripComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            PropertyInfo field = ((ColorComboBoxItem)colorToolStripComboBox.SelectedItem).Field;
            ImageListViewColor color = (ImageListViewColor)field.GetValue(null, null);
            imageListView1.Colors = color;
        }

        private void detailsToolStripButton_Click(object sender, EventArgs e)
        {
            imageListView1.View = View.Details;
        }

        private void thumbnailsToolStripButton_Click(object sender, EventArgs e)
        {
            imageListView1.View = View.Thumbnails;
        }

        private void galleryToolStripButton_Click(object sender, EventArgs e)
        {
            imageListView1.View = View.Gallery;
        }

        private void paneToolStripButton_Click(object sender, EventArgs e)
        {
            imageListView1.View = View.Pane;
        }

        private void clearThumbsToolStripButton_Click(object sender, EventArgs e)
        {
            imageListView1.ClearThumbnailCache();
        }

        private void x96ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.ThumbnailSize = new Size(96, 96);
        }

        private void x120ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.ThumbnailSize = new Size(120, 120);
        }

        private void x200ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.ThumbnailSize = new Size(200, 200);
        }

        private void showCheckboxesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.ShowCheckBoxes = !imageListView1.ShowCheckBoxes;
        }

        private void showFileIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.ShowFileIcons = !imageListView1.ShowFileIcons;
        }

        private void allowCheckBoxClickToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowCheckBoxClick = !imageListView1.AllowCheckBoxClick;
        }

        private void allowColumnClickToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowColumnClick = !imageListView1.AllowColumnClick;
        }

        private void allowColumnResizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowColumnResize = !imageListView1.AllowColumnResize;
        }

        private void allowPaneResizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowPaneResize = !imageListView1.AllowPaneResize;
        }

        private void multiSelectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.MultiSelect = !imageListView1.MultiSelect;
        }

        private void allowDragToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowDrag = !imageListView1.AllowDrag;
        }

        private void allowDropToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowDrop = !imageListView1.AllowDrop;
        }

        private void allowDuplicateFilenamesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.AllowDuplicateFileNames = !imageListView1.AllowDuplicateFileNames;
        }

        private void continuousCacheModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (imageListView1.CacheMode == CacheMode.Continuous)
                imageListView1.CacheMode = CacheMode.OnDemand;
            else
                imageListView1.CacheMode = CacheMode.Continuous;
        }

        private void integralScrollToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.IntegralScroll = !imageListView1.IntegralScroll;
        }

        private void imageListView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if ((e.Buttons & MouseButtons.Right) != MouseButtons.None)
            {
                // Group menu
                for (int j = groupByToolStripMenuItem.DropDownItems.Count - 1; j >= 0; j--)
                {
                    if (groupByToolStripMenuItem.DropDownItems[j].Tag != null)
                        groupByToolStripMenuItem.DropDownItems.RemoveAt(j);
                }
                int i = 0;
                foreach (ImageListView.ImageListViewColumnHeader col in imageListView1.Columns)
                {
                    ToolStripMenuItem item = new ToolStripMenuItem(col.Text);
                    item.Checked = (imageListView1.GroupColumn == i);
                    item.Tag = i;
                    item.Click += new EventHandler(groupColumnMenuItem_Click);
                    groupByToolStripMenuItem.DropDownItems.Insert(i, item);
                    i++;
                }
                if (i == 0)
                {
                    ToolStripMenuItem item = new ToolStripMenuItem("None");
                    item.Enabled = false;
                    groupByToolStripMenuItem.DropDownItems.Insert(0, item);
                }

                // Sort menu
                for (int j = sortByToolStripMenuItem.DropDownItems.Count - 1; j >= 0; j--)
                {
                    if (sortByToolStripMenuItem.DropDownItems[j].Tag != null)
                        sortByToolStripMenuItem.DropDownItems.RemoveAt(j);
                }
                i = 0;
                foreach (ImageListView.ImageListViewColumnHeader col in imageListView1.Columns)
                {
                    ToolStripMenuItem item = new ToolStripMenuItem(col.Text);
                    item.Checked = (imageListView1.SortColumn == i);
                    item.Tag = i;
                    item.Click += new EventHandler(sortColumnMenuItem_Click);
                    sortByToolStripMenuItem.DropDownItems.Insert(i, item);
                    i++;
                }
                if (i == 0)
                {
                    ToolStripMenuItem item = new ToolStripMenuItem("None");
                    item.Enabled = false;
                    sortByToolStripMenuItem.DropDownItems.Insert(0, item);
                }

                // Show menu
                columnContextMenu.Show(imageListView1, e.Location);
            }
        }

        private void groupColumnMenuItem_Click(object sender, EventArgs e)
        {
            int i = (int)((ToolStripMenuItem)sender).Tag;
            imageListView1.GroupColumn = i;
        }

        private void sortColumnMenuItem_Click(object sender, EventArgs e)
        {
            int i = (int)((ToolStripMenuItem)sender).Tag;
            imageListView1.SortColumn = i;
        }

        private void groupAscendingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.GroupOrder = SortOrder.Ascending;
        }

        private void sortAscendingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.SortOrder = SortOrder.Ascending;
        }

        private void groupDescendingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.GroupOrder = SortOrder.Descending;
        }

        private void sortDescendingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imageListView1.SortOrder = SortOrder.Descending;
        }

        #endregion

        #region Change Selection/Checkboxes
        private void imageListView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                if (e.KeyCode == Keys.A)
                    imageListView1.SelectAll();
                else if (e.KeyCode == Keys.U)
                    imageListView1.ClearSelection();
                else if (e.KeyCode == Keys.I)
                    imageListView1.InvertSelection();
            }
            else if (e.Alt)
            {
                if (e.KeyCode == Keys.A)
                    imageListView1.CheckAll();
                else if (e.KeyCode == Keys.U)
                    imageListView1.UncheckAll();
                else if (e.KeyCode == Keys.I)
                    imageListView1.InvertCheckState();
            }
        }
        #endregion

        private void imageListView1_SelectionChanged(object sender, EventArgs e)
        {
            ImageListViewItem sel = null;
            bool any = (imageListView1.SelectedItems.Count > 0);

            btnDelete.Enabled = any;
            btnRename.Enabled = false; // TODO rename NYI any;
            if (!any)
                return;

            sel = imageListView1.SelectedItems[0];
        }

        private string _zippath;

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archive Files|*.zip;*.7z;*.cbr;*.cbz;*.rar";
            if (!string.IsNullOrEmpty(LastFile))
                ofd.InitialDirectory = Path.GetDirectoryName(LastFile);
            ofd.Multiselect = false;
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            LastFile = ofd.FileName;
            var filename = ofd.FileName;
            try
            {
                Cursor = Cursors.WaitCursor;
                BuildListViewZip(filename); // TODO background?
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO clean up temp dirs ?
            Close();
        }

        private void BuildListViewZip(string zippath)
        {
            // 1. path to archive file
            // 2. if not supported, done
            // 3. open with 7z. if fail, done
            // 4. clear list view
            // 5. scan archive file.
            //  a. supported images: add as virtual item, get a thumbnail
            //  b. other files: what to do?

            if (!InitArchive())
                return;

            using (SevenZipExtractor extr = new SevenZipExtractor(zippath))
            {

                IReadOnlyCollection<ArchiveFileInfo> zipentries;
                // 2. Get the list of archive entries [if fail, done]
                try
                {
                    zipentries = extr.ArchiveFileData;
                    if (zipentries.Count == 0)
                    {
                        MessageBox.Show("Empty/invalid archive");
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("Empty/invalid archive");
                    return;
                }

                var adapt = new ArchiveAdaptor(zippath);

                List<string> filelist = new List<string>(zipentries.Count);
                foreach (var entry in zipentries)
                {
                    if (entry.IsDirectory)
                        continue;
                    var extension = Path.GetExtension(entry.FileName).ToLower();
                    if (IsImage(extension) || IsText(extension))
                    {
                        filelist.Add(entry.FileName);
                    }
                    // TODO other files - esp. zip of zips
                    // TODO consider adding all files to filelist; have adaptor provide placeholder
                }

                var arr = filelist.ToArray();
                Array.Sort(arr, new WindowsNaturalSort());

                imageListView1.Items.Clear();
                imageListView1.SuspendLayout();

                foreach (var fp in arr)
                {
                    var fn = Path.GetFileName(fp);
                    var extension = Path.GetExtension(fp).ToLower();

                    ImageListViewItem item = new ImageListViewItem((object)fp, fn);
                    imageListView1.Items.Add(item, adapt);
                }

                imageListView1.Items[0].Selected = true;

                imageListView1.ResumeLayout();
                _zippath = zippath;
            }
        }

        private bool InitArchive()
        {
            try
            {
                // Need to invoke 32 or 64 bit DLL as required for running OS (_not_ the build target!)
                if (Environment.Is64BitOperatingSystem)
                    SevenZipExtractor.SetLibraryPath(Path.Combine(Application.StartupPath, "7z-64.dll"));
                else
                    SevenZipExtractor.SetLibraryPath(Path.Combine(Application.StartupPath, "7z-32.dll"));
            }
            catch
            {
                MessageBox.Show("7z dll missing");
                return false;
            }

            return true;
        }

        private string[] images = { ".jpg", ".jpeg", ".gif", ".png", ".bmp" };
        private string[] text = { ".txt" };
        private bool IsImage(string ext)
        {
            return isFoo(ext, images);
        }
        private bool IsText(string ext)
        {
            return isFoo(ext, text);
        }

        private bool isFoo(string ext, string [] list)
        { 
            var ext2 = ext.ToLower();
            return list.FirstOrDefault(x => x == ext2) != null;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (imageListView1.SelectedItems.Count < 1)
                return;

            //imageListView1.Items.Remove(imageListView1.SelectedItems[0]);

            // Remove the selected item AND set selection to the next item.
            var tofind = imageListView1.SelectedItems[0];
            int i = 0;
            foreach (var item in imageListView1.Items)
            {
                if (item == tofind)
                    break;
                i++;
            }
            imageListView1.Items.RemoveAt(i);
            if (i >= imageListView1.Items.Count)
                i = imageListView1.Items.Count - 1;
            imageListView1.Items[i].Selected = true;
            imageListView1.Items[i].Focused = true;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sfd = new SaveFileDialog();
            sfd.FileName = Path.GetFileNameWithoutExtension(_zippath) + "_out_.zip";
            sfd.InitialDirectory = Path.GetDirectoryName(_zippath);
            sfd.OverwritePrompt = true;
            if (DialogResult.OK == sfd.ShowDialog())
            {
                WriteZip(sfd.FileName);
            }
        }

        private int targetheight = 1400; // TODO consider GUI mechanism?
        private static long targetJpegQuality = 75L; // TODO consider GUI mechanism?

        private void WriteZip(string outpath)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string tempdir = Path.Combine(Path.GetTempPath(), "MangaEdit");
                if (Directory.Exists(tempdir))
                    Directory.Delete(tempdir, true);
                Directory.CreateDirectory(tempdir);

                using (SevenZipExtractor extr = new SevenZipExtractor(_zippath))
                {
                    foreach (var item in imageListView1.Items)
                    {
                        string itempath = (string)item.VirtualItemKey;
                        string outpath2 = Path.Combine(tempdir, itempath);
                        Directory.CreateDirectory(Path.GetDirectoryName(outpath2));

                        // For resize:
                        // 1. if image.Hight > targetHigh
                        // a. Use adaptor.GetThumbnail() to get image of size(65535,targetHigh)
                        // b. SaveJpeg(targetfile, image)
                        if (item.Dimensions.Height > targetheight)
                        {
                            using (var newimg = item.Adaptor.GetThumbnail(item.VirtualItemKey,
                                new Size(65535, targetheight),
                                UseEmbeddedThumbnails.Never, false))
                            {
                                SaveJpeg(outpath2, newimg);
                            }
                        }
                        else
                        {
                            // No resize but make sure save to jpeg
                            using (var newimg = item.Adaptor.GetThumbnail(item.VirtualItemKey,
                                new Size(65535, 65535),
                                UseEmbeddedThumbnails.Never, false))
                            {
                                SaveJpeg(outpath2, newimg);
                            }
                        }
                    }
                }

                var comp = new SevenZipCompressor();
                comp.ArchiveFormat = OutArchiveFormat.Zip;
                comp.DirectoryStructure = true;
                comp.CompressDirectory(tempdir, outpath, recursion: true);

                Directory.Delete(tempdir, true);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }


        public static void SaveJpeg(string path, Image image)
        {
            SaveJpeg(path, image, targetJpegQuality);
        }
        public static void SaveJpeg(string path, Image image, long quality)
        {
            using (var encoderParameters = new System.Drawing.Imaging.EncoderParameters(1))
            using (var encoderParameter = new System.Drawing.Imaging.EncoderParameter(Encoder.Quality, quality))
            {
                ImageCodecInfo codecInfo = ImageCodecInfo.GetImageDecoders().First(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
                encoderParameters.Param[0] = encoderParameter;
                image.Save(path, codecInfo, encoderParameters);
            }
        }
        private void DemoForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private string LastFile { get; set; }

        private void LoadSettings()
        {
            var _mysettings = DASettings.Load();

            // No existing settings. Use default.
            if (_mysettings.Fake)
            {
                StartPosition = FormStartPosition.CenterScreen;
            }
            else
            {
                // restore windows position
                StartPosition = FormStartPosition.Manual;
                Top = _mysettings.WinTop;
                Left = _mysettings.WinLeft;
                Height = _mysettings.WinHigh;
                Width = _mysettings.WinWide;
                LastFile = _mysettings.LastPath;

                imageListView1.PaneWidth = _mysettings.SplitLoc;
                imageListView1.ThumbnailSize = new Size(_mysettings.Thumbsize,
                    _mysettings.Thumbsize);
            }

        }
        private void SaveSettings()
        {
            var _mysettings = new DASettings();
            var bounds = DesktopBounds;
            _mysettings.WinTop = Location.Y;
            _mysettings.WinLeft = Location.X;
            _mysettings.WinHigh = Size.Height;
            _mysettings.WinWide = Size.Width;
            _mysettings.Fake = false;
            _mysettings.LastPath = LastFile;
            //_mysettings.PathHistory = mnuMRU.GetFiles().ToList();
            _mysettings.SplitLoc = imageListView1.PaneWidth;
            _mysettings.Thumbsize = imageListView1.ThumbnailSize.Height;
            _mysettings.Save();
        }
    }

    public class WindowsNaturalSort : IComparer<string>
    {
        [System.Runtime.InteropServices.DllImport("shlwapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(String x, String y);

        public int Compare(string filePath1, string filePath2)
        {
            //var basename1 = Path.GetFileName(filePath1);
            //var basename2 = Path.GetFileName(filePath2);

            //return StrCmpLogicalW(basename1, basename2);
            return StrCmpLogicalW(filePath1, filePath2);
        }

    }

    public class ArchiveAdaptor : ImageListViewItemAdaptors.FileSystemAdaptor
    {
        private SevenZipExtractor _extr;
        private string _zippath;

        public ArchiveAdaptor(string zippath)
        {
            _zippath = zippath;
        }

        private Image UnpackImage(string key)
        {
            try
            {
                using (SevenZipExtractor extr = new SevenZipExtractor(_zippath))
                {
                    using (MemoryStream mem = new MemoryStream())
                    {
                        extr.ExtractFile(key, mem);
                        if (Path.GetExtension(key).ToLower() == ".txt")
                        {
                            return TextToImage(mem.ToArray());
                        }
                        var ms = new MemoryStream(mem.GetBuffer());  // dispose exception if don't copy stream
                        return Image.FromStream(ms);
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public override Image GetThumbnail(object key, Size size, UseEmbeddedThumbnails useEmbeddedThumbnails, bool useExifOrientation)
        {
            if (disposed)
                return null;

            Image img = UnpackImage((string)key);
            if (img == null)
                return null;
            Image ret = Extractor.Instance.GetThumbnail(img, size, useEmbeddedThumbnails, useExifOrientation);
            img.Dispose();
            return ret;
        }

        public override Utility.Tuple<ColumnType, string, object>[] GetDetails(object key)
        {
            if (disposed)
                return null;

            Image img = UnpackImage((string)key);
            List<Utility.Tuple<ColumnType, string, object>> details = new List<Utility.Tuple<ColumnType, string, object>>();

            // Get file info
            if (img != null)
            {
                //FileInfo info = new FileInfo(filename);
                //details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.DateCreated, string.Empty, info.CreationTime));
                //details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.DateAccessed, string.Empty, info.LastAccessTime));
                //details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.DateModified, string.Empty, info.LastWriteTime));
                //details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.FileSize, string.Empty, info.Length));
                //details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.FilePath, string.Empty, info.DirectoryName ?? ""));
                //details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.FolderName, string.Empty, info.Directory.Name ?? ""));

                // Get metadata
                Metadata metadata = Extractor.Instance.GetMetadata(img);
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.Dimensions, string.Empty, new Size(metadata.Width, metadata.Height)));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.Resolution, string.Empty, new SizeF((float)metadata.DPIX, (float)metadata.DPIY)));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.ImageDescription, string.Empty, metadata.ImageDescription ?? ""));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.EquipmentModel, string.Empty, metadata.EquipmentModel ?? ""));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.DateTaken, string.Empty, metadata.DateTaken));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.Artist, string.Empty, metadata.Artist ?? ""));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.Copyright, string.Empty, metadata.Copyright ?? ""));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.ExposureTime, string.Empty, (float)metadata.ExposureTime));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.FNumber, string.Empty, (float)metadata.FNumber));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.ISOSpeed, string.Empty, (ushort)metadata.ISOSpeed));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.UserComment, string.Empty, metadata.Comment ?? ""));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.Rating, string.Empty, (ushort)metadata.Rating));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.Software, string.Empty, metadata.Software ?? ""));
                details.Add(new Utility.Tuple<ColumnType, string, object>(ColumnType.FocalLength, string.Empty, (float)metadata.FocalLength));

                img.Dispose();
            }

            return details.ToArray();
        }


        private Image TextToImage(byte[] arr)
        {
            var text = System.Text.Encoding.Default.GetString(arr);
            using (Bitmap bmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                using (Font f = new Font("Arial", 18))
                {
                    SizeF stringSz = g.MeasureString(text, f);
                    Bitmap bmp2 = new Bitmap(bmp, (int)Math.Ceiling(stringSz.Width),
                                                  (int)Math.Ceiling(stringSz.Height));
                    using (Graphics g2 = Graphics.FromImage(bmp2))
                    {
                        //g2.FillRectangle(Brushes.White, new RectangleF(0,0,bmp2.Size.Width,bmp2.Size.Height));
                        g2.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                        g2.DrawString(text, f, Brushes.Black, 0, 0);
                    }
                    return bmp2;
                }
            }
        }

    }
}

// TODO imagelistview splitter pos
// TODO window size & pos
// TODO option to save to other file formats
// TODO display image dimensions on select
// TODO display ctrl+click to disable multi-select in imagelistview?
