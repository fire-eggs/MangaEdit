using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Net;
using SevenZip;

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
            imageListView1.Columns.Add(ColumnType.Name);
            imageListView1.Columns.Add(ColumnType.Dimensions);
            imageListView1.Columns.Add(ColumnType.FileSize);
            imageListView1.Columns.Add(ColumnType.FolderName);

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
            }
            toolStripStatusLabel1.Text = string.Format("{0} Items: {1} Selected{3}, {2} Checked",
                imageListView1.Items.Count, imageListView1.SelectedItems.Count, imageListView1.CheckedItems.Count,
                selname);

            groupAscendingToolStripMenuItem.Checked = imageListView1.GroupOrder == SortOrder.Ascending;
            groupDescendingToolStripMenuItem.Checked = imageListView1.GroupOrder == SortOrder.Descending;
            sortAscendingToolStripMenuItem.Checked = imageListView1.SortOrder == SortOrder.Ascending;
            sortDescendingToolStripMenuItem.Checked = imageListView1.SortOrder == SortOrder.Descending;
        }
        #endregion

        #region Set ImageListView options
        private void label1_Click(object sender, EventArgs e)
        {
            if (ofBrowseImage.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in ofBrowseImage.FileNames)
                {
                    imageListView1.Items.Add(file);
                }
            }
        }

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

        #region Set selected image to PropertyGrid
        private void imageListView1_SelectionChanged(object sender, EventArgs e)
        {
            ImageListViewItem sel = null;
            bool any = (imageListView1.SelectedItems.Count > 0);

            button1.Enabled = any;
            button2.Enabled = any;
            label1.Text = "";
            if (!any)
                return;

            sel = imageListView1.SelectedItems[0];
            label1.Text = sel.Text;
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

        #region Update folder list asynchronously
        private void PopulateListView(DirectoryInfo path)
        {
            imageListView1.Items.Clear();
            imageListView1.SuspendLayout();
            int i = 0;
            foreach (FileInfo p in path.GetFiles("*.*"))
            {
                if (p.Name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".cur", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".emf", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".wmf", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".tif", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase) ||
                    p.Name.EndsWith(".gif", StringComparison.OrdinalIgnoreCase))
                {
                    imageListView1.Items.Add(p.FullName);
                    if (i == 1) imageListView1.Items[imageListView1.Items.Count - 1].Enabled = false;
                    i++;
                    if (i == 3) i = 0;
                }
            }
            imageListView1.ResumeLayout();
        }

        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            KeyValuePair<TreeNode, List<TreeNode>> kv = (KeyValuePair<TreeNode, List<TreeNode>>)e.Result;
            TreeNode rootNode = kv.Key;
            List<TreeNode> nodes = kv.Value;
            if (rootNode.Tag == null)
            {
                //treeView1.Nodes.Clear();
                //foreach (TreeNode node in nodes)
                //    treeView1.Nodes.Add(node);
            }
            else
            {
                KeyValuePair<DirectoryInfo, bool> ktag = (KeyValuePair<DirectoryInfo, bool>)rootNode.Tag;
                rootNode.Tag = new KeyValuePair<DirectoryInfo, bool>(ktag.Key, true);
                rootNode.Nodes.Clear();
                foreach (TreeNode node in nodes)
                    rootNode.Nodes.Add(node);
            }
        }

        private static void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            TreeNode rootNode = e.Argument as TreeNode;

            List<TreeNode> nodes = GetNodes(rootNode);

            e.Result = new KeyValuePair<TreeNode, List<TreeNode>>(rootNode, nodes);
        }

        private static List<TreeNode> GetNodes(TreeNode rootNode)
        {
            if (rootNode.Tag == null)
            {
                List<TreeNode> volNodes = new List<TreeNode>();
                foreach (DriveInfo info in System.IO.DriveInfo.GetDrives())
                {
                    if (info.IsReady && info.DriveType == DriveType.Fixed)
                    {
                        DirectoryInfo rootPath = info.RootDirectory;
                        TreeNode volNode = new TreeNode(info.VolumeLabel + " (" + info.Name + ")", 0, 0);
                        volNode.Tag = new KeyValuePair<DirectoryInfo, bool>(rootPath, false);
                        List<TreeNode> nodes = GetNodes(volNode);
                        volNode.Tag = new KeyValuePair<DirectoryInfo, bool>(rootPath, true);
                        volNode.Nodes.Clear();
                        foreach (TreeNode node in nodes)
                            volNode.Nodes.Add(node);

                        volNode.Expand();
                        volNodes.Add(volNode);
                    }
                }

                return volNodes;
            }
            else
            {
                KeyValuePair<DirectoryInfo, bool> kv = (KeyValuePair<DirectoryInfo, bool>)rootNode.Tag;
                bool done = kv.Value;
                if (done)
                    return new List<TreeNode>();

                DirectoryInfo rootPath = kv.Key;
                List<TreeNode> nodes = new List<TreeNode>();

                DirectoryInfo[] dirs = new DirectoryInfo[0];
                try
                {
                    dirs = rootPath.GetDirectories();
                }
                catch
                {
                    return new List<TreeNode>();
                }
                foreach (DirectoryInfo info in dirs)
                {
                    if ((info.Attributes & FileAttributes.System) != FileAttributes.System)
                    {
                        TreeNode aNode = new TreeNode(info.Name, 1, 2);
                        aNode.Tag = new KeyValuePair<DirectoryInfo, bool>(info, false);
                        GetDirectories(aNode);
                        nodes.Add(aNode);
                    }
                }
                return nodes;
            }
        }

        private static void GetDirectories(TreeNode node)
        {
            KeyValuePair<DirectoryInfo, bool> ktag = (KeyValuePair<DirectoryInfo, bool>)node.Tag;
            DirectoryInfo rootPath = ktag.Key;

            DirectoryInfo[] dirs = new DirectoryInfo[0];
            try
            {
                dirs = rootPath.GetDirectories();
            }
            catch
            {
                return;
            }
            foreach (DirectoryInfo info in dirs)
            {
                if ((info.Attributes & FileAttributes.System) != FileAttributes.System)
                {
                    TreeNode aNode = new TreeNode(info.Name, 1, 2);
                    aNode.Tag = new KeyValuePair<DirectoryInfo, bool>(info, false);
                    if (GetDirCount(info) != 0)
                    {
                        aNode.Nodes.Add("Dummy1");
                    }
                    node.Nodes.Add(aNode);
                }
            }
            node.Tag = new KeyValuePair<DirectoryInfo, bool>(ktag.Key, true);
        }

        private static int GetDirCount(DirectoryInfo rootPath)
        {
            DirectoryInfo[] dirs = new DirectoryInfo[0];
            try
            {
                dirs = rootPath.GetDirectories();
            }
            catch
            {
                return 0;
            }

            return dirs.Length;
        }
        #endregion

        private string _zippath;

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archive Files|*.zip;*.7z;*.cbr;*.cbz;*.rar";
            // ofd.InitialDirectory // TODO remember & restore
            ofd.Multiselect = false;
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            var filename = ofd.FileName;
            BuildListViewZip(filename); // TODO background?
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // TODO clean up temp dirs 
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
            foreach (var img in list)
                if (ext2 == img)
                    return true;
            return false;
        }

        private Image TextToImage(byte [] arr)
        {
            var text = System.Text.Encoding.Default.GetString(arr);
            using (Bitmap bmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                using (Font f = new Font("Courier New", 14))
                {
                    SizeF stringSz = g.MeasureString(text, f);
                    Bitmap bmp2 = new Bitmap(bmp, (int)Math.Ceiling(stringSz.Width), 
                                                  (int)Math.Ceiling(stringSz.Height));
                    using (Graphics g2 = Graphics.FromImage(bmp2))
                    {
                        g.DrawString(text, f, Brushes.Black, 0, 0);
                    }
                    return bmp2;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (imageListView1.SelectedItems.Count < 1)
                return;
            imageListView1.Items.Remove(imageListView1.SelectedItems[0]);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sfd = new SaveFileDialog();
            sfd.FileName = Path.GetFileNameWithoutExtension(_zippath) + "_out_.zip";
            //sfd.FileName = Path.GetFileName(_zippath);
            sfd.InitialDirectory = Path.GetDirectoryName(_zippath);
            sfd.OverwritePrompt = true;
            if (DialogResult.OK == sfd.ShowDialog())
            {
                WriteZip(sfd.FileName);
            }
        }

        private void WriteZip(string outpath)
        {
            MessageBox.Show(outpath);

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
                    using (FileStream fs = new FileStream(outpath2, FileMode.Create))
                    {
                        extr.ExtractFile(itempath, fs);
                    }
                }
            }

            var comp = new SevenZipCompressor();
            comp.ArchiveFormat = OutArchiveFormat.Zip;
            comp.DirectoryStructure = true;
            comp.CompressDirectory(tempdir, outpath, recursion: true);
        }
    }

    public class WindowsNaturalSort : IComparer<string>
    {
        [System.Runtime.InteropServices.DllImport("shlwapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(String x, String y);

        public int Compare(string filePath1, string filePath2)
        {
            var basename1 = Path.GetFileName(filePath1);
            var basename2 = Path.GetFileName(filePath2);

            return StrCmpLogicalW(basename1, basename2);
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

        public override Image GetThumbnail(object key, Size size, UseEmbeddedThumbnails useEmbeddedThumbnails, bool useExifOrientation)
        {
            if (disposed)
                return null;

            try
            {
                using (SevenZipExtractor extr = new SevenZipExtractor(_zippath))
                {
                    using (MemoryStream mem = new MemoryStream())
                    {
                        extr.ExtractFile((string)key, mem);
                        if (Path.GetExtension((string)key).ToLower() == ".txt")
                        {
                            using (Image img = TextToImage(mem.ToArray()))
                            {
                                return Extractor.Instance.GetThumbnail(img, size, useEmbeddedThumbnails, useExifOrientation);
                            }
                        }
                        using (Image img = Image.FromStream(mem))
                        {
                            return Extractor.Instance.GetThumbnail(img, size, useEmbeddedThumbnails, useExifOrientation);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
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
