using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using Spire.Xls;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Text.Json;

namespace Toppling
{
    public partial class MainForm : Form
    {
        #region Form Program

        private float scale = 1.0f;
        private PointF offset = new PointF(0, 0);
        private const int GRID_SIZE = 20;
        private const float POINT_RADIUS = 0.02f;
        private ToolTip toolTip = new ToolTip();
        private PointF? hoveredPoint = null;
        private StatusStrip statusStrip = new StatusStrip();
        private ToolStripStatusLabel coordinatesLabel = new ToolStripStatusLabel();
        private List<BlockXY> blockCoords = new List<BlockXY>();
        private List<PointF> waterLevelPoints = new List<PointF>();
        private List<PointF> slopePoints = new List<PointF>();


        private void InitializeStatusStrip()
        {
            try
            {
                // Ensure the creation and configuration of the status bar
                if (statusStrip == null)
                {
                    statusStrip = new StatusStrip();
                }

                if (statusStrip.IsDisposed)
                {
                    statusStrip = new StatusStrip();
                }

                if (coordinatesLabel == null)
                {
                    coordinatesLabel = new ToolStripStatusLabel();
                }

                if (coordinatesLabel.IsDisposed)
                {
                    coordinatesLabel = new ToolStripStatusLabel();
                }

                // Clear and re-add items
                statusStrip.Items.Clear();

                // Configure coordinate labels
                coordinatesLabel.Text = "coordinate: (0.00, 0.00)";
                coordinatesLabel.AutoSize = true;
                coordinatesLabel.BorderSides = ToolStripStatusLabelBorderSides.All;
                coordinatesLabel.BorderStyle = Border3DStyle.SunkenOuter;

                // Add to status bar
                statusStrip.Items.Add(coordinatesLabel);

                // Configure status bar
                statusStrip.Dock = DockStyle.Bottom;
                statusStrip.SizingGrip = false;

                // Ensure the status bar has been added to the form
                if (!this.IsDisposed && !this.Controls.Contains(statusStrip))
                {
                    this.Controls.Add(statusStrip);
                }

                // Set the correct position
                if (!this.IsDisposed)
                {
                    statusStrip.Location = new Point(0, this.ClientSize.Height - statusStrip.Height);
                    statusStrip.Width = this.ClientSize.Width;
                    statusStrip.Visible = true;
                    statusStrip.BringToFront();
                }
            }
            catch (ObjectDisposedException)
            {
                // Ignore exceptions for disposed objects
            }
            catch (Exception ex)
            {
                // Catch and ignore any exceptions to avoid affecting application stability
                Console.WriteLine($"Error initializing status bar: {ex.Message}");
            }
        }

        private void MainForm_Resize(object? sender, EventArgs e)
        {
            // Ensure status bar is at the bottom when form size changes
            if (statusStrip != null)
            {
                statusStrip.Visible = true;
                statusStrip.BringToFront();
            }
        }

        private void MainForm_Load(object? sender, EventArgs e)
        {
            try
            {
                // Initialize parameters
                _IP = new GlobalParameter();
                initext_load();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Initialize PictureBox
                if (pictureBox1 != null && !pictureBox1.IsDisposed && pictureBox1.Image == null)
                {
                    pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

                    // Draw initial image
                    using (Graphics g = Graphics.FromImage(pictureBox1.Image))
                    {
                        g.Clear(Color.White);
                        DrawGrid(g);
                    }
                    pictureBox1.Invalidate();
                }

                // Safely initialize controls
                if (!this.IsDisposed)
                {
                    // Ensure all controls are visible - but don't force in Load event, only in Shown event

                    // Layout adjustments needed
                    this.PerformLayout();
                }
            }
            catch (ObjectDisposedException)
            {
                // Ignore exceptions for disposed objects
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred in Load event: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MainForm_Shown(object? sender, EventArgs e)
        {
            try
            {
                if (this.IsDisposed) return;

                // Ensure all controls are visible after form is shown
                EnsureAllControlsAreVisible();

                // Ensure status bar is shown
                if (!this.IsDisposed)
                {
                    InitializeStatusStrip();

                    // Immediately update the interface
                    Application.DoEvents();

                    // Ensure key controls are visible
                    if (pictureBox1 != null && !pictureBox1.IsDisposed)
                    {
                        pictureBox1.Visible = true;
                        pictureBox1.Invalidate(); // Redraw
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // Ignore exceptions for disposed objects
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred in Shown event: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PictureBox_MouseLeave(object? sender, EventArgs e)
        {
            toolTip.Hide(pictureBox1);
            hoveredPoint = null;
            if (coordinatesLabel != null)
            {
                coordinatesLabel.Text = "coordinate: (0.00, 0.00)";
            }
        }

        private void PictureBox_MouseMove(object? sender, MouseEventArgs e)
        {
            UpdateCoordinates(e.Location);

            if (isDragging)
            {
                offset.X += (e.X - lastMousePosition.X) / scale;
                offset.Y -= (e.Y - lastMousePosition.Y) / scale;
                lastMousePosition = e.Location;
                pictureBox1.Invalidate();
                return;
            }

            // Check if mouse is near a point
            PointF? nearestPoint = FindNearestPoint(e.Location);
            if (nearestPoint.HasValue)
            {
                if (!hoveredPoint.HasValue || nearestPoint.Value != hoveredPoint.Value)
                {
                    hoveredPoint = nearestPoint.Value;
                    toolTip.Show($"coordinate: ({nearestPoint.Value.X:F2}, {nearestPoint.Value.Y:F2})",
                               pictureBox1,
                               e.Location.X + 20,
                               e.Location.Y - 20);
                    pictureBox1.Invalidate(); // Redraw to show highlighted point
                }
            }
            else if (hoveredPoint.HasValue)
            {
                toolTip.Hide(pictureBox1);
                hoveredPoint = null;
                pictureBox1.Invalidate(); // Redraw to remove highlighted point
            }
        }

        private void UpdateCoordinates(Point mousePos)
        {
            // Adjust mouse coordinate calculation to match the coordinate system used for drawing
            float mouseX = (mousePos.X - offset.X * scale) / scale;
            float mouseY = (pictureBox1.Height - mousePos.Y - offset.Y * scale) / scale;
            if (coordinatesLabel != null)
            {
                coordinatesLabel.Text = $"coordinate: ({mouseX:F2}, {mouseY:F2})";
            }
        }

        private PointF? FindNearestPoint(Point mousePos)
        {
            // Adjust mouse coordinate calculation
            float mouseX = (mousePos.X - offset.X * scale) / scale;
            float mouseY = (pictureBox1.Height - mousePos.Y - offset.Y * scale) / scale;

            PointF? nearestPoint = null;
            float minDistance = float.MaxValue;

            // Check all quadrilateral vertices
            foreach (var block in blockCoords)
            {
                PointF[] quadPoints = new PointF[]
                {
                    new PointF((float)block.xTL, (float)block.yTL),
                    new PointF((float)block.xBL, (float)block.yBL),
                    new PointF((float)block.xBR, (float)block.yBR),
                    new PointF((float)block.xTR, (float)block.yTR)
                };

                foreach (var point in quadPoints)
                {
                    float distance = (float)Math.Sqrt(
                        Math.Pow(point.X - mouseX, 2) +
                        Math.Pow(point.Y - mouseY, 2));

                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        nearestPoint = point;
                    }
                }
            }

            // Check water level points
            foreach (var point in waterLevelPoints)
            {
                float distance = (float)Math.Sqrt(
                    Math.Pow(point.X - mouseX, 2) +
                    Math.Pow(point.Y - mouseY, 2));

                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearestPoint = point;
                }
            }

            // Check slope face points
            foreach (var point in slopePoints)
            {
                float distance = (float)Math.Sqrt(
                    Math.Pow(point.X - mouseX, 2) +
                    Math.Pow(point.Y - mouseY, 2));

                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearestPoint = point;
                }
            }

            // Increase detection range by using a larger POINT_RADIUS
            if (minDistance <= POINT_RADIUS * 10)  // Increased detection range
            {
                return nearestPoint;
            }

            return null;
        }

        private bool isDragging = false;
        private Point lastMousePosition;

        private void PictureBox_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                lastMousePosition = e.Location;
                toolTip.Hide(pictureBox1);
            }
        }

        private void PictureBox_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = false;
            }
        }

        private void PictureBox_MouseWheel(object? sender, MouseEventArgs e)
        {
            Point mousePos = pictureBox1.PointToClient(MousePosition);
            float mouseX = (mousePos.X - offset.X * scale) / scale;
            float mouseY = (mousePos.Y - offset.Y * scale) / scale;

            float oldScale = scale;
            // Use a smaller scaling factor for smoother zooming
            float scaleFactor = 1.1f;
            if (e.Delta > 0)
                scale *= scaleFactor;
            else
                scale /= scaleFactor;

            // Limit zoom range
            scale = Math.Max(0.1f, Math.Min(scale, 100.0f));

            if (scale != oldScale)
            {
                // Keep mouse position unchanged
                offset.X = mousePos.X / scale - mouseX;
                offset.Y = mousePos.Y / scale - mouseY;
                pictureBox1.Invalidate();
                UpdateCoordinates(mousePos);
            }
        }

        private void DrawGrid(Graphics g)
        {
            // Draw grid
            using (Pen gridPen = new Pen(Color.LightGray, 1))
            {
                gridPen.DashStyle = DashStyle.Dot;

                int horizontalLines = pictureBox1.Height / GRID_SIZE;
                int verticalLines = pictureBox1.Width / GRID_SIZE;

                for (int i = 0; i <= horizontalLines; i++)
                {
                    int y = i * GRID_SIZE;
                    g.DrawLine(gridPen, 0, y, pictureBox1.Width, y);
                }

                for (int i = 0; i <= verticalLines; i++)
                {
                    int x = i * GRID_SIZE;
                    g.DrawLine(gridPen, x, 0, x, pictureBox1.Height);
                }
            }
        }

        private void DrawWaterLevel(Graphics g)
        {
            if (waterLevelPoints.Count < 2) return;

            using (Pen waterPen = new Pen(Color.Blue, 0.03f))
            {
                waterPen.Alignment = PenAlignment.Center;
                waterPen.DashStyle = DashStyle.Dash;

                // Draw water level line using multiple segments
                PointF[] points = waterLevelPoints.ToArray();
                g.DrawLines(waterPen, points);
            }
        }

        private void DrawSlopeFace(Graphics g)
        {
            if (slopePoints.Count < 2) return;

            using (Pen slopePen = new Pen(Color.Gray, 0.05f))
            {
                slopePen.Alignment = PenAlignment.Center;

                // Draw slope face polyline
                PointF[] points = slopePoints.ToArray();
                g.DrawLines(slopePen, points);
            }
        }

        private void DrawPoints(Graphics g)
        {
            // Draw all quadrilateral vertices
            foreach (var block in blockCoords)
            {
                PointF[] quadPoints = new PointF[]
                {
                    new PointF((float)block.xTL, (float)block.yTL),
                    new PointF((float)block.xBL, (float)block.yBL),
                    new PointF((float)block.xBR, (float)block.yBR),
                    new PointF((float)block.xTR, (float)block.yTR)
                };

                foreach (var point in quadPoints)
                {
                    using (SolidBrush brush = new SolidBrush(Color.Red))
                    {
                        float radius = (point == hoveredPoint) ? POINT_RADIUS * 2 : POINT_RADIUS;
                        g.FillEllipse(brush,
                            point.X - radius,
                            point.Y - radius,
                            radius * 2,
                            radius * 2);
                    }
                }
            }

            // Draw water level points
            foreach (var point in waterLevelPoints)
            {
                using (SolidBrush brush = new SolidBrush(Color.Blue))
                {
                    float radius = (point == hoveredPoint) ? POINT_RADIUS * 2 : POINT_RADIUS;
                    g.FillEllipse(brush,
                        point.X - radius,
                        point.Y - radius,
                        radius * 2,
                        radius * 2);
                }
            }

            // Draw slope face points
            foreach (var point in slopePoints)
            {
                using (SolidBrush brush = new SolidBrush(Color.Green))
                {
                    float radius = (point == hoveredPoint) ? POINT_RADIUS * 2 : POINT_RADIUS;
                    g.FillEllipse(brush,
                        point.X - radius,
                        point.Y - radius,
                        radius * 2,
                        radius * 2);
                }
            }
        }

        private void pictureBox1_Paint(object? sender, PaintEventArgs e)
        {
            if (pictureBox1.Image == null)
            {
                pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            }

            using (Graphics g = Graphics.FromImage(pictureBox1.Image))
            {
                g.Clear(Color.White);
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;

                DrawGrid(g);

                GraphicsState state = g.Save();

                // Adjust coordinate system
                g.TranslateTransform(0, pictureBox1.Height);  // Move to bottom left corner
                g.ScaleTransform(scale, -scale);  // Invert Y-axis direction
                g.TranslateTransform(offset.X, offset.Y);

                // Draw slope face
                DrawSlopeFace(g);

                // Draw water level
                DrawWaterLevel(g);

                // Draw all quadrilaterals
                foreach (var block in blockCoords)
                {
                    PointF[] quadPoints = new PointF[]
                    {
                        new PointF((float)block.xTL, (float)block.yTL),
                        new PointF((float)block.xBL, (float)block.yBL),
                        new PointF((float)block.xBR, (float)block.yBR),
                        new PointF((float)block.xTR, (float)block.yTR)
                    };

                    using (Pen pen = new Pen(Color.Red, 0.01f))
                    {
                        pen.Alignment = PenAlignment.Center;
                        // Draw only three sides (left, bottom, right)
                        g.DrawLine(pen, quadPoints[0], quadPoints[1]); // Left side
                        g.DrawLine(pen, quadPoints[1], quadPoints[2]); // Bottom side
                        g.DrawLine(pen, quadPoints[2], quadPoints[3]); // Right side
                    }
                }

                // Draw points (including highlighted points)
                DrawPoints(g);

                g.Restore(state);
            }

            // Draw Image to PictureBox
            e.Graphics.DrawImage(pictureBox1.Image, 0, 0);
        }

        public void UpdateBlockCoordinates(List<BlockXY> newCoords)
        {
            blockCoords.Clear();
            blockCoords.AddRange(newCoords);
        }

        public void UpdateWaterLevel(List<PointF> points)
        {
            waterLevelPoints.Clear();
            waterLevelPoints.AddRange(points);
        }

        public void UpdateSlopePoints(List<PointF> points)
        {
            slopePoints.Clear();
            slopePoints.AddRange(points);

            // Calculate bounds of all points
            double minX = double.MaxValue;
            double maxX = double.MinValue;
            double minY = double.MaxValue;
            double maxY = double.MinValue;

            // Calculate bounds of quadrilaterals
            foreach (var block in blockCoords)
            {
                minX = Math.Min(minX, Math.Min(block.xTL, Math.Min(block.xBL, Math.Min(block.xBR, block.xTR))));
                maxX = Math.Max(maxX, Math.Max(block.xTL, Math.Max(block.xBL, Math.Max(block.xBR, block.xTR))));
                minY = Math.Min(minY, Math.Min(block.yTL, Math.Min(block.yBL, Math.Min(block.yBR, block.yTR))));
                maxY = Math.Max(maxY, Math.Max(block.yTL, Math.Max(block.yBL, Math.Max(block.yBR, block.yTR))));
            }

            // Calculate bounds of water level points
            foreach (var point in waterLevelPoints)
            {
                minX = Math.Min(minX, point.X);
                maxX = Math.Max(maxX, point.X);
                minY = Math.Min(minY, point.Y);
                maxY = Math.Max(maxY, point.Y);
            }

            // Calculate bounds of slope face points
            foreach (var point in slopePoints)
            {
                minX = Math.Min(minX, point.X);
                maxX = Math.Max(maxX, point.X);
                minY = Math.Min(minY, point.Y);
                maxY = Math.Max(maxY, point.Y);
            }

            // Calculate center point
            double centerX = (minX + maxX) / 2;
            double centerY = (minY + maxY) / 2;

            // Calculate width and height
            double width = maxX - minX;
            double height = maxY - minY;

            // Calculate appropriate scale
            float scaleX = (float)(pictureBox1.Width / (width * 1.2));
            float scaleY = (float)(pictureBox1.Height / (height * 1.2));
            scale = Math.Min(scaleX, scaleY);

            // Calculate offset to center the image
            offset.X = (float)(pictureBox1.Width / (2 * scale) - centerX);
            offset.Y = (float)(pictureBox1.Height / (2 * scale) - centerY);

        }

        public void UpdateAllData(List<BlockXY> blocks, List<PointF> waterPoints, List<PointF> slopePoints)
        {
            UpdateBlockCoordinates(blocks);
            UpdateWaterLevel(waterPoints);
            UpdateSlopePoints(slopePoints);

            // Trigger a single redraw after all data is updated
            pictureBox1.Invalidate();
        }

        private GlobalParameter _IP = null!;
        private string templatePath = null!;
        private string resultPath = null!;
        private string resultsPath = null!;
        private ExcelPackage package = null!;
        private Workbook workbook = null!;


        public MainForm()
        {
            // First, perform basic initialization
            InitializeComponent();

            // Initialize status bar, but don't add it to control collection yet
            InitializeStatusStrip();

            // Register Load and Shown events, which will be triggered when the form is ready
            this.Load += MainForm_Load;
            this.Shown += MainForm_Shown;

            // Initialize tooltip
            toolTip.AutoPopDelay = 5000;
            toolTip.InitialDelay = 100;
            toolTip.ReshowDelay = 100;
            toolTip.ShowAlways = true;

            // Register event handlers after all controls have been created
            if (pictureBox1 != null && !pictureBox1.IsDisposed)
            {
                pictureBox1.MouseWheel += PictureBox_MouseWheel;
                pictureBox1.MouseDown += PictureBox_MouseDown;
                pictureBox1.MouseMove += PictureBox_MouseMove;
                pictureBox1.MouseUp += PictureBox_MouseUp;
                pictureBox1.MouseLeave += PictureBox_MouseLeave;
            }
        }

        // Add a method to ensure all controls are visible
        private void EnsureAllControlsAreVisible()
        {
            try
            {
                // Force show all controls
                this.SuspendLayout();

                // Traverse all controls and ensure they are visible
                RecursiveSetControlVisible(this);

                // Handle some important control types separately
                Control[] controls = new Control[this.Controls.Count];
                this.Controls.CopyTo(controls, 0);

                foreach (Control c in controls)
                {
                    if (c == null || c.IsDisposed) continue;

                    if (c is TableLayoutPanel tablePanel)
                    {
                        tablePanel.Visible = true;
                        Control[] panelControls = new Control[tablePanel.Controls.Count];
                        tablePanel.Controls.CopyTo(panelControls, 0);

                        foreach (Control child in panelControls)
                        {
                            if (child != null && !child.IsDisposed)
                            {
                                child.Visible = true;
                            }
                        }
                    }
                }

                // Explicitly set visibility and Z-order for some important controls
                if (button1 != null && !button1.IsDisposed)
                {
                    button1.Visible = true;
                    button1.BringToFront();
                }

                if (pictureBox1 != null && !pictureBox1.IsDisposed)
                {
                    pictureBox1.Visible = true;
                    // Ensure PictureBox is at the bottom
                    pictureBox1.SendToBack();
                }

                // Ensure status bar is at the top
                if (statusStrip != null && !statusStrip.IsDisposed)
                {
                    // First, check if status bar has already been added to the form
                    if (!this.Controls.Contains(statusStrip))
                    {
                        this.Controls.Add(statusStrip);
                    }

                    statusStrip.Visible = true;
                    statusStrip.BringToFront();

                    // Ensure coordinate label is visible
                    if (coordinatesLabel != null && !coordinatesLabel.IsDisposed && !statusStrip.Items.Contains(coordinatesLabel))
                    {
                        statusStrip.Items.Add(coordinatesLabel);
                    }

                    // Force refresh status bar
                    statusStrip.Refresh();
                }

                // Force relayout and refresh
                this.ResumeLayout(true);
                this.PerformLayout();
                this.Refresh();
            }
            catch (ObjectDisposedException)
            {
                // Ignore exception for disposed objects
                return;
            }
        }

        // Recursively set control and its child controls visible
        private void RecursiveSetControlVisible(Control control)
        {
            if (control == null || control.IsDisposed) return;

            try
            {
                // Set current control visible
                control.Visible = true;

                // Create temporary list to avoid exception due to collection being modified
                Control[] childControls = new Control[control.Controls.Count];
                control.Controls.CopyTo(childControls, 0);

                // Perform the same operation for each child control recursively
                foreach (Control child in childControls)
                {
                    if (child != null && !child.IsDisposed)
                    {
                        RecursiveSetControlVisible(child);
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // Ignore exception for disposed objects and continue
                return;
            }
        }
        #endregion

        #region Main Functions
        private void button1_Click(object? sender, EventArgs e)
        {
            loadInterfaceParameters();
            if (ValidateInput(_IP.MeanDipA, _IP.MeanDipB, _IP.MeanDipDirA, _IP.MeanDipDirB, _IP.MeanSpaceA, _IP.MeanSpaceB, _IP.MeanFricA, _IP.MeanFricB, _IP.SlopeHeight, _IP.SlopeAngle, _IP.TopAngle, _IP.SlopeDipDir,
                _IP.MeanSeis, _IP.PorePress, _IP.UnitWeight, _IP.UnitWeightH2O, _IP.FisherKA, _IP.FisherKB, _IP.StDevSpaceA, _IP.StDevSpaceB, _IP.StDevFricA, _IP.StDevFricB, _IP.StDevSeis, _IP.StDevPorePress, _IP.StDevUnitWt,
                _IP.StDevUnitWtH20, _IP.NumTrials, _IP.DistSpaceA, _IP.DistSpaceB, _IP.DistFricA, _IP.DistFricB, _IP.DistSeis, _IP.DistPorePress, _IP.DistUnitWeight, _IP.DistUnitWeightH2O))
            {
                Setfilepaths();
                CheckFileIntegrity();
                PreviewMeanGeometry();
            }
            else
            {
                MessageBox.Show("Invalid input parameters, please check the input values.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object? sender, EventArgs e)
        {
            if (IfAddSupport())
            {
                MonteCarlo();
                CalcProb();
                UpdateHistogram();
                OpenExcelFile();
            }
        }
        private void CheckFileIntegrity()
        {
            if (package.Workbook.Worksheets["Analysis Input"] == null)
            {
                MessageBox.Show("Template file does not contain 'Analysis Input' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (package.Workbook.Worksheets["Add Support"] == null)
            {
                MessageBox.Show("Template file does not contain 'Add Support' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (package.Workbook.Worksheets["Results"] == null)
            {
                MessageBox.Show("Template file does not contain 'Results' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (package.Workbook.Worksheets["Frequency"] == null)
            {
                MessageBox.Show("Template file does not contain 'Frequency' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (package.Workbook.Worksheets["Analysis Details 1"] == null)
            {
                MessageBox.Show("Template file does not contain 'Analysis Details 1' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (package.Workbook.Worksheets["Analysis Details 2"] == null)
            {
                MessageBox.Show("Template file does not contain 'Analysis Details 2' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (package.Workbook.Worksheets["Analysis Details 3"] == null)
            {
                MessageBox.Show("Template file does not contain 'Analysis Details 3' worksheet", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void saveImage(string sheetName, string imageName, int[] ints)
        {
            workbook = new Workbook();
            workbook.LoadFromFile(resultPath);

            Worksheet sheet = workbook.Worksheets[sheetName];

            string imagePath = Path.Combine(resultsPath, imageName + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".png");

            sheet.SaveToImage(imagePath, ints[0], ints[1], ints[2], ints[3]);
            workbook.Dispose();
            try
            {
                Process.Start(new System.Diagnostics.ProcessStartInfo(imagePath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred while opening png file: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool IfAddSupport()
        {
            if (checkBox1.Checked || checkBox2.Checked)
            {
                // Get MeanDipA and MeanFricA values
                double meanDipA = double.Parse(textBox2_11.Text);
                double meanFricA = double.Parse(textBox3_31.Text);

                AddSupportForm form3;

                // If no parameters are imported (Magnitude is 0), use simple constructor
                if (_IP.Magnitude == 0)
                {
                    form3 = new AddSupportForm(meanDipA, meanFricA);
                }
                else
                {
                    form3 = new AddSupportForm(
                        _IP.NatureOfForceApplication,
                        _IP.Magnitude,
                        _IP.Orientation * 180 / Math.PI,  // Convert to degrees
                        _IP.OptimumOrientationAgainstSliding * 180 / Math.PI,  // Convert to degrees
                        _IP.OptimumOrientationAgainstToppling * 180 / Math.PI,  // Convert to degrees
                        _IP.EffectiveWidth,
                        meanDipA,  // Pass MeanDipA
                        meanFricA  // Pass MeanFricA
                    );
                }

                using (form3)
                {
                    if (form3.ShowDialog() == DialogResult.OK)
                    {
                        // Get additional parameters from Form3
                        _IP.NatureOfForceApplication = form3.NatureOfForceApplication;
                        _IP.Magnitude = form3.Magnitude;
                        _IP.Orientation = form3.Orientation;  // Already converted to radians in AddSupportForm
                        _IP.OptimumOrientationAgainstSliding = form3.OptimumOrientationAgainstSliding;  // Already converted to radians in AddSupportForm
                        _IP.OptimumOrientationAgainstToppling = form3.OptimumOrientationAgainstToppling;  // Already converted to radians in AddSupportForm
                        _IP.EffectiveWidth = form3.EffectiveWidth;
                        return true;
                    }
                    return false;
                }
            }
            return true;
        }

        private void Setfilepaths()
        {
            // Use EPPlus to write Excel
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string rootPath = Path.GetFullPath(System.IO.Path.Combine(exePath, @"..\..\..\..\..\"));
            string resourcePath = Path.Combine(rootPath, "Resources");
            resultsPath = Path.Combine(rootPath, "Results");
            templatePath = Path.Combine(resourcePath, "template.xlsx");
            resultPath = Path.Combine(resultsPath, "result_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            // Copy template file
            File.Copy(templatePath, resultPath, true);
            package = new(new FileInfo(resultPath));
        }

        private void loadInterfaceParameters()
        {
            _IP.SlopeHeight = double.Parse(textBox1_1.Text);
            _IP.SlopeAngle = double.Parse(textBox1_2.Text) * (Math.PI / 180);
            _IP.TopAngle = double.Parse(textBox1_3.Text) * (Math.PI / 180);
            _IP.SlopeDipDir = double.Parse(textBox1_4.Text) * (Math.PI / 180);

            _IP.MeanDipA = double.Parse(textBox2_11.Text) * (Math.PI / 180);
            _IP.MeanDipB = double.Parse(textBox2_21.Text) * (Math.PI / 180);
            _IP.MeanDipDirA = double.Parse(textBox2_12.Text) * (Math.PI / 180);
            _IP.MeanDipDirB = double.Parse(textBox2_22.Text) * (Math.PI / 180);
            _IP.MeanSpaceA = double.Parse(textBox3_11.Text);
            _IP.MeanSpaceB = double.Parse(textBox3_21.Text);
            _IP.MeanFricA = double.Parse(textBox3_31.Text) * (Math.PI / 180);
            _IP.MeanFricB = double.Parse(textBox3_41.Text) * (Math.PI / 180);

            _IP.MeanSeis = double.Parse(textBox4_11.Text);
            _IP.PorePress = textBox4_21.Text.EndsWith('%') == true ? double.Parse(textBox4_21.Text.TrimEnd('%')) * 0.01 : double.Parse(textBox4_21.Text);
            _IP.UnitWeight = double.Parse(textBox3_51.Text);
            _IP.UnitWeightH2O = double.Parse(textBox3_61.Text);

            _IP.DistSpaceA = (comboBox1.Text);
            _IP.DistSpaceB = (comboBox2.Text);
            _IP.DistFricA = (comboBox3.Text);
            _IP.DistFricB = (comboBox4.Text);
            _IP.DistUnitWeight = (comboBox5.Text);
            _IP.DistUnitWeightH2O = (comboBox6.Text);
            _IP.DistSeis = (comboBox7.Text);
            _IP.DistPorePress = (comboBox8.Text);

            _IP.FisherKA = double.Parse(textBox2_13.Text);
            _IP.FisherKB = double.Parse(textBox2_23.Text);
            _IP.StDevSpaceA = double.Parse(textBox3_13.Text);
            _IP.StDevSpaceB = double.Parse(textBox3_23.Text);
            _IP.StDevFricA = double.Parse(textBox3_33.Text) * (Math.PI / 180);
            _IP.StDevFricB = double.Parse(textBox3_43.Text) * (Math.PI / 180);
            _IP.StDevSeis = double.Parse(textBox4_13.Text);
            _IP.StDevPorePress = textBox4_23.Text.EndsWith('%') == true ? double.Parse(textBox4_23.Text.TrimEnd('%')) * 0.01 : double.Parse(textBox4_23.Text);
            _IP.StDevUnitWt = double.Parse(textBox3_53.Text);
            _IP.StDevUnitWtH20 = double.Parse(textBox3_63.Text);
            _IP.BoltBlocksTogether = checkBox2.Checked;
            _IP.AddToeSupport = checkBox1.Checked;

            _IP.NumTrials = int.Parse(textBox5.Text);

        }
        private void initext_load()
        {
            richTextBox1.Text = "Slope Height, H (m)";
            richTextBox1.Select(14, 1);
            richTextBox1.SelectionFont = new Font("Times New Roman", 11, FontStyle.Italic);

            richTextBox2.Text = "Slope Angle, ψs (°)";
            // Set "ψs" as italic
            richTextBox2.Select(13, 2); // Select "ψs"
            richTextBox2.SelectionFont = new Font("Times New Roman", 13, FontStyle.Italic);
            // Optional: Display "s" as subscript
            richTextBox2.Select(14, 1); // Select "s"
            richTextBox2.SelectionCharOffset = -5; // Vertical offset (simulate subscript)

            richTextBox3.Text = "Top Angle, ψts (°)";
            richTextBox3.Select(11, 3);
            richTextBox3.SelectionFont = new Font("Times New Roman", 13, FontStyle.Italic);
            richTextBox3.Select(12, 2);
            richTextBox3.SelectionCharOffset = -5;

            richTextBox4.Text = "Dip Direction of Slope, as (°)";
            richTextBox4.Select(24, 2);
            richTextBox4.SelectionFont = new Font("Times New Roman", 11, FontStyle.Italic);
            richTextBox4.Select(25, 1);
            richTextBox4.SelectionCharOffset = -5;

            richTextBox5.Text = "Set A (base plane set), ψa, αa";
            richTextBox5.Select(24, 2);
            richTextBox5.SelectionFont = new Font("Times New Roman", 12, FontStyle.Italic);
            richTextBox5.Select(28, 2);
            richTextBox5.SelectionFont = new Font("Times New Roman", 12, FontStyle.Italic);
            richTextBox5.Select(25, 1);
            richTextBox5.SelectionCharOffset = -6;
            richTextBox5.Select(29, 1);
            richTextBox5.SelectionCharOffset = -6;

            richTextBox6.Text = "Set B (sub-vertical set), ψb, αb";
            richTextBox6.Select(26, 2);
            richTextBox6.SelectionFont = new Font("Times New Roman", 12, FontStyle.Italic);
            richTextBox6.Select(30, 2);
            richTextBox6.SelectionFont = new Font("Times New Roman", 12, FontStyle.Italic);
            richTextBox6.Select(27, 1);
            richTextBox6.SelectionCharOffset = -7;
            richTextBox6.Select(31, 1);
            richTextBox6.SelectionCharOffset = -7;

            richTextBox7.Text = ("Spacing Set A (m), Sa");
            richTextBox7.Select(19, 2);
            richTextBox7.SelectionFont = new Font("Times New Roman", 11, FontStyle.Italic);
            richTextBox7.Select(20, 1);
            richTextBox7.SelectionCharOffset = -5;

            richTextBox8.Text = ("Spacing Set B (m), Sb");
            richTextBox8.Select(19, 2);
            richTextBox8.SelectionFont = new Font("Times New Roman", 11, FontStyle.Italic);
            richTextBox8.Select(20, 1);
            richTextBox8.SelectionCharOffset = -5;

            richTextBox9.Text = ("Friction Angle of Set A (°), φa");
            richTextBox9.Select(29, 2);
            richTextBox9.SelectionFont = new Font("Times New Roman", 13, FontStyle.Italic);
            richTextBox9.Select(30, 1);
            richTextBox9.SelectionCharOffset = -5;

            richTextBox10.Text = ("Friction Angle of Set B (°), φb");
            richTextBox10.Select(29, 2);
            richTextBox10.SelectionFont = new Font("Times New Roman", 13, FontStyle.Italic);
            richTextBox10.Select(30, 1);
            richTextBox10.SelectionCharOffset = -7;

            richTextBox11.Text = ("Unit Weight of Rock (kN/m3), γrock");
            richTextBox11.Select(29, 5);
            richTextBox11.SelectionFont = new Font("Times New Roman", 11, FontStyle.Italic);
            richTextBox11.Select(30, 4);
            richTextBox11.SelectionCharOffset = -6;
            richTextBox11.Select(24, 2);
            richTextBox11.SelectedText = "m³";

            richTextBox12.Text = ("Unit Weight of Water (kN/m3), γwater");
            richTextBox12.Select(30, 6);
            richTextBox12.SelectionFont = new Font("Times New Roman", 11, FontStyle.Italic);
            richTextBox12.Select(31, 5);
            richTextBox12.SelectionCharOffset = -6;
            richTextBox12.Select(25, 2);
            richTextBox12.SelectedText = "m³";
        }

        private void ifPreviewGeometry()
        {
            DialogResult reply = MessageBox.Show(
                "Input data is valid. Do you want to preview the mean slope geometry?",
                "Preview Geometry",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information
            );
            if (reply == DialogResult.Cancel)
            {
                MessageBox.Show("Analysis stopped", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0);
            }
        }

        private void ifPerformAnalysis()
        {
            DialogResult reply = MessageBox.Show(
                "Monte Carlo simulation has been completed. Do you want to check the results?",
                "Perform Analysis",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );
            if (reply == DialogResult.No)
            {
                MessageBox.Show("Analysis stopped", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0);
            }
            else
            {
                package.Save();
                saveImage("Results", "Results", new int[4] { 2, 1, 28, 21 });
            }
        }

        private static bool ValidateInput(double MeanDipA, double MeanDipB, double MeanDipDirA, double MeanDipDirB, double MeanSpaceA, double MeanSpaceB, double MeanFricA, double MeanFricB, double SlopeHeight, double SlopeAngle, double TopAngle, double SlopeDipDir
            , double MeanSeis, double PorePress, double UnitWeight, double UnitWeightH2O, double FisherKA, double FisherKB, double StDevSpaceA, double StDevSpaceB, double StDevFricA, double StDevFricB, double StDevSeis, double StDevPorePress, double StDevUnitWt,
            double StDevUnitWtH20, double NumTrials, string DistSpaceA, string DistSpaceB, string DistFricA, string DistFricB, string DistSeis, string DistPorePress, string DistUnitWeight, string DistUnitWeightH2O)
        {
            bool isValid = true;

            if (MeanDipA <= 0 || MeanDipA > 90)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Dip of Set A", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanDipB <= 0 || MeanDipB > 90)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Dip of Set B", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanFricA <= 0 || MeanFricA > 90)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Friction Angle of Set A", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanFricB <= 0 || MeanFricB > 90)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Friction Angle of Set B", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (SlopeAngle <= 0 || SlopeAngle > 90)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Slope Angle", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanDipDirA < 0 || MeanDipDirA > 360)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Dip Direction of Set A", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanDipDirB < 0 || MeanDipDirB > 360)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Dip Direction of Set B", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (SlopeDipDir < 0 || SlopeDipDir > 360)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Dip Direction of Slope", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanSpaceA <= 0)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Spacing of Set A", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanSpaceB <= 0)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Spacing of Set B", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (SlopeHeight <= 0)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Slope Height", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (UnitWeight <= 0)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Unit Weight", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (TopAngle < 0 || TopAngle > SlopeAngle)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Top Angle", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (UnitWeightH2O <= 0 || UnitWeightH2O > 11.75)
            {
                isValid = false;
                MessageBox.Show("Invalid Value For Unit Weight of Water", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (PorePress < 0 || PorePress > 1)
            {
                isValid = false;
                MessageBox.Show("Pore Pressure Entry Must Be Between 0% and 100%", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (NumTrials <= 0)
            {
                isValid = false;
                MessageBox.Show("Number of Trials must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (MeanDipA > SlopeAngle)
            {
                isValid = false;
                MessageBox.Show("Mean dip of set A should be less than the slope angle", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (FisherKA <= 0 || FisherKB <= 0)
            {
                DialogResult Reply = MessageBox.Show("Fisher constant must be > 0 for discontinuity to be randomly sampled. Do you wish to continue with the analysis? If yes, the discontinuity orientation will be treated as a fixed value."
                    , "Input Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (Reply == DialogResult.No)
                {
                    isValid = false;
                    return false;
                }
            }

            if ((DistSpaceA == "Lognormal" || DistSpaceA == "Exponential") && MeanSpaceA <= 0)
            {
                MessageBox.Show("The mean value of Spacing Set A must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistSpaceA != "Fixed" && DistSpaceA != "Exponential" && StDevSpaceA <= 0)
            {
                MessageBox.Show("The Standard Deviation of Spacing Set A must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistSpaceB == "Lognormal" || DistSpaceB == "Exponential") && MeanSpaceB <= 0)
            {
                MessageBox.Show("The mean value of Spacing Set B must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistSpaceB != "Fixed" && DistSpaceB != "Exponential" && StDevSpaceB <= 0)
            {
                MessageBox.Show("The Standard Deviation of Spacing Set B must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistFricA == "Lognormal" || DistFricA == "Exponential") && MeanFricA <= 0)
            {
                MessageBox.Show("The mean value of Friction Angle of Set A must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistFricA != "Fixed" && DistFricA != "Exponential" && StDevFricA <= 0)
            {
                MessageBox.Show("The Standard Deviation of Friction Angle of Set A must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistFricB == "Lognormal" || DistFricB == "Exponential") && MeanFricB <= 0)
            {
                MessageBox.Show("The mean value of Friction Angle of Set B must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistFricB != "Fixed" && DistFricB != "Exponential" && StDevFricB <= 0)
            {
                MessageBox.Show("The Standard Deviation of Friction Angle of Set B must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistUnitWeight == "Lognormal" || DistUnitWeight == "Exponential") && UnitWeight <= 0)
            {
                MessageBox.Show("The mean value of Unit Weight of Rock must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistUnitWeight != "Fixed" && DistUnitWeight != "Exponential" && StDevUnitWt <= 0)
            {
                MessageBox.Show("The Standard Deviation of Unit Weight of Rock must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistUnitWeightH2O == "Lognormal" || DistUnitWeightH2O == "Exponential") && UnitWeightH2O <= 0)
            {
                MessageBox.Show("The mean value of Unit Weight of Water must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistUnitWeightH2O != "Fixed" && DistUnitWeightH2O != "Exponential" && StDevUnitWtH20 <= 0)
            {
                MessageBox.Show("The Standard Deviation of Unit Weight of Water must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistSeis == "Lognormal" || DistSeis == "Exponential") && MeanSeis <= 0)
            {
                MessageBox.Show("The mean value of Seismic Coefficient must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistSeis != "Fixed" && DistSeis != "Exponential" && StDevSeis <= 0)
            {
                MessageBox.Show("The Standard Deviation of Seismic Coefficient must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if ((DistPorePress == "Lognormal" || DistPorePress == "Exponential") && PorePress <= 0)
            {
                MessageBox.Show("The mean value of Pore Pressure must be > 0 if using Lognormal or Exponential distributions", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }
            if (DistPorePress != "Fixed" && DistPorePress != "Exponential" && StDevPorePress <= 0)
            {
                MessageBox.Show("The Standard Deviation of Pore Pressure must be > 0", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isValid = false;
                return false;
            }

            double PlungeB = Math.PI / 2 - MeanDipB;
            double TrendB = MeanDipDirB + Math.PI;
            bool MeanDipDirAOK, TrendBOK, PlungeBOK;
            if (TrendB > 2 * Math.PI)
            {
                TrendB -= 2 * Math.PI;
            }

            // Check Set A and Set B trends and plunges
            double delta = DegreeToRadian(40);
            if (SlopeDipDir + delta > 2 * Math.PI)
            {
                if (MeanDipDirA > SlopeDipDir + delta - 2 * Math.PI && MeanDipDirA < SlopeDipDir - delta)
                    MeanDipDirAOK = false;
                else
                    MeanDipDirAOK = true;
                if (TrendB > SlopeDipDir + delta - 2 * Math.PI && TrendB < SlopeDipDir - delta)
                    TrendBOK = false;
                else
                    TrendBOK = true;
            }
            else if (SlopeDipDir - delta < 0)
            {
                if (MeanDipDirA < SlopeDipDir - delta + 2 * Math.PI && MeanDipDirA > SlopeDipDir + delta)
                    MeanDipDirAOK = false;
                else
                    MeanDipDirAOK = true;
                if (TrendB < SlopeDipDir - delta + 2 * Math.PI && TrendB > SlopeDipDir + delta)
                    TrendBOK = false;
                else
                    TrendBOK = true;
            }
            else
            {
                if (MeanDipDirA < SlopeDipDir - delta || MeanDipDirA > SlopeDipDir + delta)
                    MeanDipDirAOK = false;
                else
                    MeanDipDirAOK = true;
                if (TrendB < SlopeDipDir - delta || TrendB > SlopeDipDir + delta)
                    TrendBOK = false;
                else
                    TrendBOK = true;
            }

            // Check Set B plunge
            if (PlungeB > SlopeAngle)
                PlungeBOK = false;
            else
                PlungeBOK = true;

            if (!MeanDipDirAOK && (!TrendBOK || !PlungeBOK))
            {
                MessageBox.Show("The specified discontinuity orientations of both set A and set B are well outside the kinematic limits for block toppling. Check input values for errors and retry.",
                    "Invalid Input for Block Toppling", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Analysis stopped", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0);
            }
            else if (MeanDipDirAOK && (!TrendBOK || !PlungeBOK))
            {
                MessageBox.Show("The specified discontinuity orientations are well outside the kinematic limits for block toppling. However, stability against planar sliding on Set A should be investigated via a sliding limit equilibrium analysis.",
                    "Invalid Input for Block Toppling", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Analysis stopped", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0);
            }
            else if (!MeanDipDirAOK && (TrendBOK || PlungeBOK))
            {
                MessageBox.Show("The specified discontinuity orientation for set B satisfies the kinematic limits for toppling. However, sliding on set A is not kinematically feasible. Therefore, block toppling is unlikely but flexural toppling remains possible.",
                    "Invalid Input for Block Toppling", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Analysis stopped", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0);
            }

            return isValid;
        }
        private static double DegreeToRadian(double degree)
        {
            return degree * Math.PI / 180.0;
        }

        private void PreviewMeanGeometry()
        {
            // Define arrays to store coordinates
            List<BlockXY> blockCoords = new List<BlockXY>();
            List<PorePressXY> waterLevel = new List<PorePressXY>();

            double appDipA, appDipB;
            int blockCount = 0;

            // Build overall geometry shape
            double toeX = 0;
            double toeY = 0;
            double crestX = _IP.SlopeHeight / Math.Tan(_IP.SlopeAngle);
            double crestY = _IP.SlopeHeight;
            double leftLimX = -10;
            double leftLimY = toeY;
            double rightLimX;

            // Define the toe as the starting point
            double botLeftX = toeX;
            double botLeftY = toeY;
            double topLeftX = toeX;
            double topLeftY = toeY;

            // Define the toe water condition
            double pwLeftX = 0;
            double pwLeftY = 0;

            appDipA = Math.Atan(Math.Tan(_IP.MeanDipA) * Math.Abs(Math.Sin(_IP.SlopeDipDir - _IP.MeanDipDirA + Math.PI / 2)));
            appDipB = Math.Atan(Math.Tan(_IP.MeanDipB) * Math.Abs(Math.Sin(_IP.SlopeDipDir - _IP.MeanDipDirB + Math.PI / 2)));

            // Iteratively generate mean geometry
            do
            {
                // Calculate the right coordinate of the current block
                double botRightY = botLeftY + (_IP.MeanSpaceB / Math.Cos(appDipA - (Math.PI / 2 - appDipB))) * Math.Sin(appDipA);
                double botRightX = botLeftX + (_IP.MeanSpaceB / Math.Cos(appDipA - (Math.PI / 2 - appDipB))) * Math.Cos(appDipA);
                double topRightY = topLeftY + (_IP.MeanSpaceB / Math.Cos(_IP.SlopeAngle - (Math.PI / 2 - appDipB))) * Math.Sin(_IP.SlopeAngle);
                double topRightX = topLeftX + (_IP.MeanSpaceB / Math.Cos(_IP.SlopeAngle - (Math.PI / 2 - appDipB))) * Math.Cos(_IP.SlopeAngle);

                double pwRightX, pwRightY;
                double hypDist, xDist;

                // Determine the current block's position relative to the crest to calculate the right coordinate
                if (topRightY < crestY && topRightX < crestX)
                {
                    // Block is below the crest, so keep the coordinates unchanged
                    pwRightX = botRightX - _IP.PorePress * (botRightX - topRightX);
                    pwRightY = botRightY + _IP.PorePress * (topRightY - botRightY);
                }
                else if (topLeftX < crestX && topRightX > crestX)
                {
                    // Block spans the crest, recalculate the right coordinate as needed
                    hypDist = Math.Sqrt(Math.Pow(crestX - topLeftX, 2) + Math.Pow(crestY - topLeftY, 2));
                    xDist = hypDist * Math.Cos(_IP.SlopeAngle - (Math.PI / 2 - appDipB));

                    topRightX = crestX + ((_IP.MeanSpaceB - xDist) / Math.Cos(Math.PI / 2 - appDipB - _IP.TopAngle)) * Math.Cos(_IP.TopAngle);
                    topRightY = crestY + ((_IP.MeanSpaceB - xDist) / Math.Cos(Math.PI / 2 - appDipB - _IP.TopAngle)) * Math.Sin(_IP.TopAngle);

                    pwRightX = botRightX - _IP.PorePress * (botRightX - topRightX);
                    pwRightY = botRightY + _IP.PorePress * (topRightY - botRightY);
                }
                else
                {
                    // Block is above the crest, recalculate the right coordinate as needed
                    topRightX = topLeftX + (_IP.MeanSpaceB / Math.Cos(_IP.TopAngle - (Math.PI / 2 - appDipB))) * Math.Cos(_IP.TopAngle);
                    topRightY = topLeftY + (_IP.MeanSpaceB / Math.Cos(_IP.TopAngle - (Math.PI / 2 - appDipB))) * Math.Sin(_IP.TopAngle);

                    pwRightX = botRightX - _IP.PorePress * (botRightX - topRightX);
                    pwRightY = botRightY + _IP.PorePress * (topRightY - botRightY);
                }

                // Condition to end block formation at the top of the slope
                if (botRightY > topRightY || botLeftY > topLeftY)
                {
                    break;
                }

                // Add calculated coordinates to the array
                blockCoords.Add(new BlockXY
                {
                    xTL = topLeftX,
                    yTL = topLeftY,
                    xBL = botLeftX,
                    yBL = botLeftY,
                    xBR = botRightX,
                    yBR = botRightY,
                    xTR = topRightX,
                    yTR = topRightY
                });

                waterLevel.Add(new PorePressXY
                {
                    xDS = pwLeftX,
                    yDS = pwLeftY,
                    xUS = pwRightX,
                    yUS = pwRightY
                });

                // Define the left coordinate of the next block in the slope
                botLeftX = blockCoords[blockCount].xBR - (_IP.MeanSpaceA / Math.Sin(Math.PI - appDipA - appDipB) * Math.Sin(Math.PI / 2 - appDipB));
                botLeftY = blockCoords[blockCount].yBR + (_IP.MeanSpaceA / Math.Sin(Math.PI - appDipA - appDipB) * Math.Cos(Math.PI / 2 - appDipB));
                topLeftX = blockCoords[blockCount].xTR;
                topLeftY = blockCoords[blockCount].yTR;
                pwLeftX = waterLevel[blockCount].xUS;
                pwLeftY = waterLevel[blockCount].yUS;

                blockCount++;

            } while (true);

            if (blockCount == 1)
            {
                DialogResult reply = MessageBox.Show(
                    "Mean geometry forms only 1 block. Do you wish to continue analysis?",
                    "Mean Geometry",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information
                );

                if (reply == DialogResult.No)
                {
                    MessageBox.Show("Analysis stopped", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Environment.Exit(0);
                }
            }

            // Calculate rightLimX
            rightLimX = Math.Max(1.1 * blockCoords[blockCount - 1].xTR, _IP.SlopeHeight / Math.Tan(appDipA));
            double rightLimY = crestY + (rightLimX - crestX) * Math.Tan(_IP.TopAngle);

            ifPreviewGeometry();

            // Create block list
            List<BlockXY> allBlockCoords = new List<BlockXY> { };
            allBlockCoords.AddRange(blockCoords);

            // Create water level point list
            List<PointF> waterLevelPoints = new List<PointF>();
            for (int i = 0; i < waterLevel.Count; i++)
            {
                waterLevelPoints.Add(new PointF((float)waterLevel[i].xDS, (float)waterLevel[i].yDS));
                waterLevelPoints.Add(new PointF((float)waterLevel[i].xUS, (float)waterLevel[i].yUS));
            }
            waterLevelPoints.Add(new PointF((float)rightLimX, (float)rightLimY));

            // Create slope line point list
            List<PointF> SlopePoints =
            [
                new PointF((float)leftLimX, (float)leftLimY),
                    new PointF((float)toeX, (float)toeY),
                    new PointF((float)crestX, (float)crestY),
                    new PointF((float)rightLimX, (float)rightLimY),
                ];

            UpdateAllData(allBlockCoords, waterLevelPoints, SlopePoints);

            // Get worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Analysis Details 1"];
            if (worksheet == null)
            {
                MessageBox.Show("Worksheet 'Analysis Details 1' not found in template file", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Write block count
                worksheet.Cells[2, 2].Value = blockCount;

                // Write overall slope geometry coordinates, using existing format
                worksheet.Cells[5, 2].Value = leftLimX;
                worksheet.Cells[6, 2].Value = toeX;
                worksheet.Cells[7, 2].Value = crestX;
                worksheet.Cells[8, 2].Value = rightLimX;
                worksheet.Cells[5, 3].Value = leftLimY;
                worksheet.Cells[6, 3].Value = toeY;
                worksheet.Cells[7, 3].Value = crestY;
                worksheet.Cells[8, 3].Value = rightLimY;

                // Write block coordinates
                int inrow = 1;
                for (int j = 0; j < blockCount; j++)
                {
                    worksheet.Cells[inrow + 9, 1].Value = $"Block {j + 1}";
                    worksheet.Cells[inrow + 9, 2].Value = blockCoords[j].xTL;
                    worksheet.Cells[inrow + 9, 3].Value = blockCoords[j].yTL;
                    worksheet.Cells[inrow + 10, 2].Value = blockCoords[j].xBL;
                    worksheet.Cells[inrow + 10, 3].Value = blockCoords[j].yBL;
                    worksheet.Cells[inrow + 11, 2].Value = blockCoords[j].xBR;
                    worksheet.Cells[inrow + 11, 3].Value = blockCoords[j].yBR;
                    worksheet.Cells[inrow + 12, 2].Value = blockCoords[j].xTR;
                    worksheet.Cells[inrow + 12, 3].Value = blockCoords[j].yTR;
                    inrow += 4;
                }
                // Write water level data
                inrow = 1;
                if (_IP.PorePress != 0)
                {
                    worksheet.Cells[2, 8].Value = _IP.PorePress;
                    for (int j = 0; j < blockCount; j++)
                    {
                        worksheet.Cells[inrow + 9, 6].Value = "U/S";
                        worksheet.Cells[inrow + 10, 6].Value = "D/S";
                        worksheet.Cells[inrow + 9, 7].Value = waterLevel[j].xDS;
                        worksheet.Cells[inrow + 9, 8].Value = waterLevel[j].yDS;
                        worksheet.Cells[inrow + 10, 7].Value = waterLevel[j].xUS;
                        worksheet.Cells[inrow + 10, 8].Value = waterLevel[j].yUS;
                        inrow += 4;
                    }
                    worksheet.Cells[inrow + 7, 7].Value = rightLimX;
                    worksheet.Cells[inrow + 7, 8].Value = rightLimY;
                }
                else
                {
                    worksheet.Cells[2, 8].Value = $"{_IP.PorePress}% Water Table Not Plotted";
                }
                // Check and delete existing worksheet before getting Analysis Input worksheet
                worksheet = package.Workbook.Worksheets["Analysis Input"];
                if (worksheet == null)
                {
                    MessageBox.Show("Worksheet 'Analysis Input' not found in template file", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Scale preview chart axes
                int maxScale;
                if (blockCount < 2)
                {
                    maxScale = (int)Math.Ceiling(rightLimY / 10.0) * 10;
                }
                else
                {
                    maxScale = (int)Math.Ceiling(Math.Max(
                        blockCoords[blockCount - 1].xBR + 5,
                        blockCoords[blockCount - 1].yTR + 5) / 10.0) * 10;
                }
                // Set up chart
                ExcelChart chart = (ExcelChart)worksheet.Drawings[0];
                chart.XAxis.MinValue = leftLimX;
                chart.XAxis.MaxValue = maxScale;
                chart.XAxis.MajorUnit = 10;
                chart.YAxis.MinValue = leftLimY;
                chart.YAxis.MaxValue = maxScale;
                chart.YAxis.MajorUnit = 10;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to Excel file: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void MonteCarlo()
        {
            // *******************************************************************************************
            // Monte Carlo Simulation
            // *******************************************************************************************
            Random rand = new Random();

            double percentComplete;
            int ini;
            bool printIteration;

            // Random variables
            double rndSlopeAngle;
            double rndTopAngle;
            double rndHeight;
            double rndSlopeDipDir;
            double rndUnitWt = 0;
            double rndUnitWtH2O = 0;
            double rndDipA;
            double rndDipB;
            double rndDipDirA;
            double rndDipDirB;
            double rndSeis = 0;
            double rndPorePress = 0;
            double rndFricA = 0;
            double rndFricB = 0;
            double rndSpaceA = 0;
            double rndSpaceB = 0;

            // Variables for log-normal distribution (mean and standard deviation of ln(x))
            double lambdaSeis, zetaSeis;
            double lambdaPorePress, zetaPorePress;
            double lambdaFricA, zetaFricA;
            double lambdaFricB, zetaFricB;
            double lambdaSpaceA = 0, zetaSpaceA = 0;
            double lambdaSpaceB = 0, zetaSpaceB = 0;
            double lambdaUnitWt, zetaUnitWt;
            double lambdaUnitWtH2O, zetaUnitWtH2O;

            // Kinematic stability variables
            bool kineSlide, kineSlideTop, kineFlex, kineBlockTop;

            // Initialize counters to zero
            _IP.KineBlockTopCount = 0;
            _IP.KineFlexCount = 0;
            _IP.KineSlideCount = 0;
            _IP.KineSlideTopCount = 0;
            _IP.ToppleCount = 0;
            _IP.StableTopCount = 0;
            _IP.UnStableTopCount = 0;
            _IP.SlideCount = 0;
            _IP.StableSlideCount = 0;
            _IP.UnstableSlideCount = 0;
            _IP.InvalidCount = 0;
            ini = 0;
            percentComplete = 0;
            printIteration = true;

            List<MonteCarloIterationResult> allResults = new List<MonteCarloIterationResult>();
            List<int> safeNum = new List<int>();

            try
            {
                do
                {
                    ini++;

                    // Sample random values for each block based on user-defined inputs
                    // Slope angle
                    rndSlopeAngle = _IP.SlopeAngle;

                    // Top angle
                    rndTopAngle = _IP.TopAngle;

                    // Slope height
                    rndHeight = _IP.SlopeHeight;

                    // Slope dip direction
                    rndSlopeDipDir = _IP.SlopeDipDir;

                    // Direction of Set A (Fisher distribution sampling)
                    FisherSample(_IP.MeanDipA, _IP.MeanDipDirA, _IP.FisherKA, out rndDipA, out rndDipDirA);

                    // Direction of Set B (Fisher distribution sampling)
                    FisherSample(_IP.MeanDipB, _IP.MeanDipDirB, _IP.FisherKB, out rndDipB, out rndDipDirB);

                    // Random sampling of friction angle for Set A
                    if (_IP.DistFricA == "Normal")
                    {
                        do
                        {
                            rndFricA = GenerateNormalRandom(_IP.MeanFricA, _IP.StDevFricA, rand);
                        } while (rndFricA <= 0);
                    }
                    else if (_IP.DistFricA == "Lognormal")
                    {
                        zetaFricA = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevFricA, 2) / Math.Pow(_IP.MeanFricA, 2))));
                        lambdaFricA = Math.Log(_IP.MeanFricA) - 0.5 * Math.Pow(zetaFricA, 2);
                        rndFricA = Math.Exp(GenerateNormalRandom(lambdaFricA, zetaFricA, rand));
                    }
                    else if (_IP.DistFricA == "Exponential")
                    {
                        rndFricA = -_IP.MeanFricA * Math.Log(1 - rand.NextDouble());
                    }
                    else if (_IP.DistFricA == "Fixed")
                    {
                        rndFricA = _IP.MeanFricA;
                    }

                    // Random sampling of friction angle for Set B
                    if (_IP.DistFricB == "Normal")
                    {
                        do
                        {
                            rndFricB = GenerateNormalRandom(_IP.MeanFricB, _IP.StDevFricB, rand);
                        } while (rndFricB <= 0);
                    }
                    else if (_IP.DistFricB == "Lognormal")
                    {
                        zetaFricB = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevFricB, 2) / Math.Pow(_IP.MeanFricB, 2))));
                        lambdaFricB = Math.Log(_IP.MeanFricB) - 0.5 * Math.Pow(zetaFricB, 2);
                        rndFricB = Math.Exp(GenerateNormalRandom(lambdaFricB, zetaFricB, rand));
                    }
                    else if (_IP.DistFricB == "Exponential")
                    {
                        rndFricB = -_IP.MeanFricB * Math.Log(1 - rand.NextDouble());
                    }
                    else if (_IP.DistFricB == "Fixed")
                    {
                        rndFricB = _IP.MeanFricB;
                    }

                    // Rock unit weight
                    if (_IP.DistUnitWeight == "Normal")
                    {
                        do
                        {
                            rndUnitWt = GenerateNormalRandom(_IP.UnitWeight, _IP.StDevUnitWt, rand);
                        } while (rndUnitWt <= 0);
                    }
                    else if (_IP.DistUnitWeight == "Lognormal")
                    {
                        zetaUnitWt = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevUnitWt, 2) / Math.Pow(_IP.UnitWeight, 2))));
                        lambdaUnitWt = Math.Log(_IP.UnitWeight) - 0.5 * Math.Pow(zetaUnitWt, 2);
                        rndUnitWt = Math.Exp(GenerateNormalRandom(lambdaUnitWt, zetaUnitWt, rand));
                    }
                    else if (_IP.DistUnitWeight == "Exponential")
                    {
                        rndUnitWt = -_IP.UnitWeight * Math.Log(1 - rand.NextDouble());
                    }
                    else if (_IP.DistUnitWeight == "Fixed")
                    {
                        rndUnitWt = _IP.UnitWeight;
                    }

                    // Water unit weight
                    if (_IP.DistUnitWeightH2O == "Normal")
                    {
                        do
                        {
                            rndUnitWtH2O = GenerateNormalRandom(_IP.UnitWeightH2O, _IP.StDevUnitWtH20, rand);
                        } while (rndUnitWtH2O <= 0);
                    }
                    else if (_IP.DistUnitWeightH2O == "Lognormal")
                    {
                        zetaUnitWtH2O = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevUnitWtH20, 2) / Math.Pow(_IP.UnitWeightH2O, 2))));
                        lambdaUnitWtH2O = Math.Log(_IP.UnitWeightH2O) - 0.5 * Math.Pow(zetaUnitWtH2O, 2);
                        rndUnitWtH2O = Math.Exp(GenerateNormalRandom(lambdaUnitWtH2O, zetaUnitWtH2O, rand));
                    }
                    else if (_IP.DistUnitWeightH2O == "Exponential")
                    {
                        rndUnitWtH2O = -_IP.UnitWeightH2O * Math.Log(1 - rand.NextDouble());
                    }
                    else if (_IP.DistUnitWeightH2O == "Fixed")
                    {
                        rndUnitWtH2O = _IP.UnitWeightH2O;
                    }

                    // Seismic load
                    if (_IP.DistSeis == "Normal")
                    {
                        do
                        {
                            rndSeis = GenerateNormalRandom(_IP.MeanSeis, _IP.StDevSeis, rand);
                        } while (rndSeis < 0);
                    }
                    else if (_IP.DistSeis == "Lognormal")
                    {
                        zetaSeis = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevSeis, 2) / Math.Pow(_IP.MeanSeis, 2))));
                        lambdaSeis = Math.Log(_IP.MeanSeis) - 0.5 * Math.Pow(zetaSeis, 2);
                        rndSeis = Math.Exp(GenerateNormalRandom(lambdaSeis, zetaSeis, rand));
                    }
                    else if (_IP.DistSeis == "Exponential")
                    {
                        rndSeis = -_IP.MeanSeis * Math.Log(1 - rand.NextDouble());
                    }
                    else if (_IP.DistSeis == "Fixed")
                    {
                        rndSeis = _IP.MeanSeis;
                    }

                    // Pore pressure reduction
                    if (_IP.DistPorePress == "Normal")
                    {
                        do
                        {
                            rndPorePress = GenerateNormalRandom(_IP.PorePress, _IP.StDevPorePress, rand);
                        } while (rndPorePress < 0 || rndPorePress > 1);
                    }
                    else if (_IP.DistPorePress == "Lognormal")
                    {
                        zetaPorePress = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevPorePress, 2) / Math.Pow(_IP.PorePress, 2))));
                        lambdaPorePress = Math.Log(_IP.PorePress) - 0.5 * Math.Pow(zetaPorePress, 2);
                        rndPorePress = Math.Exp(GenerateNormalRandom(lambdaPorePress, zetaPorePress, rand));
                    }
                    else if (_IP.DistPorePress == "Exponential")
                    {
                        rndPorePress = -_IP.PorePress * Math.Log(1 - rand.NextDouble());
                    }
                    else if (_IP.DistPorePress == "Fixed")
                    {
                        rndPorePress = _IP.PorePress;
                    }

                    // *******************************************************************************************
                    // Check kinematic stability
                    // *******************************************************************************************
                    KinematicCheck((double)rndFricA, (double)rndDipDirA, (double)rndDipA,
                        (double)rndFricB, (double)rndDipDirB, (double)rndDipB,
                        (double)rndSlopeAngle, (double)rndSlopeDipDir,
                        out kineSlide, out kineSlideTop, out kineFlex, out kineBlockTop);

                    // *******************************************************************************************
                    // Check dynamic stability when kinematically feasible
                    // *******************************************************************************************

                    if (kineBlockTop) // Continue with limit equilibrium equation for block toppling
                    {
                        LimitEquilibrium(rndSlopeAngle, rndTopAngle, rndHeight, rndSlopeDipDir, rndDipA, rndDipB,
                            rndFricA, rndFricB, rndUnitWt, rndUnitWtH2O, rndSeis, rndPorePress,
                            ref rndSpaceA, ref rndSpaceB, ref lambdaSpaceA, ref zetaSpaceA,
                            ref lambdaSpaceB, ref zetaSpaceB, ref printIteration, ini);
                    }
                    else
                    {
                        safeNum.Add(ini);
                    }

                    allResults.Add(new MonteCarloIterationResult
                    {
                        IterationNumber = ini,
                        DipA = rndDipA,
                        DipDirA = rndDipDirA,
                        DipB = rndDipB,
                        DipDirB = rndDipDirB,
                        FricA = rndFricA,
                        FricB = rndFricB,
                        KineBlockTop = kineBlockTop,
                        KineSlide = kineSlide,
                        KineFlex = kineFlex,
                        KineSlideTop = kineSlideTop
                    });

                    percentComplete = (double)ini / _IP.NumTrials;

                } while (ini < _IP.NumTrials);

                WriteAllMonteCarloResultsToExcel(allResults, safeNum);


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during Monte Carlo simulation: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        /// <summary>
        /// Fisher distribution sampling - Generates random fracture orientations based on the method described by Fisher, Lewis and Embleton (1987)
        /// </summary>
        /// <param name="meanDip">Mean dip angle (degrees)</param>
        /// <param name="meanDipDir">Mean dip direction (degrees)</param>
        /// <param name="fisherK">Fisher constant K</param>
        /// <param name="rndDip">Output: Randomly generated dip angle (degrees)</param>
        /// <param name="rndDipDir">Output: Randomly generated dip direction (degrees)</param>
        private void FisherSample(double meanDip, double meanDipDir, double fisherK,
            out double rndDip, out double rndDipDir)
        {
            const double Pi = Math.PI;
            Random rand = new Random();

            double meanDipRad = meanDip;
            double meanDipDirRad = meanDipDir;

            if (fisherK > 0)
            {
                // Calculate plunge and trend of mean pole
                double plunge = Pi / 2 - meanDipRad;
                double trend = meanDipDirRad + Pi;
                if (trend > 2 * Pi)
                {
                    trend = trend - 2 * Pi;
                }

                // Fisher distribution parameters
                double lambda = Math.Exp(-2 * fisherK);
                double r1 = rand.NextDouble();
                double r2 = rand.NextDouble();
                double tmp = -Math.Log10(r1 * (1 - lambda) + lambda) / (2 * fisherK);
                double theta = 2 * Math.Asin(Math.Sqrt(tmp));
                theta = Pi / 2 - theta;
                double phi = 2 * Pi * r2;
                if (phi >= 2 * Pi)
                {
                    phi = phi - 2 * Pi;
                }

                // Define rotation matrix elements
                double[,] a = new double[3, 3];
                a[0, 0] = Math.Sin(plunge) * Math.Cos(trend);
                a[0, 1] = -Math.Sin(trend);
                a[0, 2] = Math.Cos(plunge) * Math.Cos(trend);
                a[1, 0] = Math.Sin(plunge) * Math.Sin(trend);
                a[1, 1] = Math.Cos(trend);
                a[1, 2] = Math.Cos(plunge) * Math.Sin(trend);
                a[2, 0] = -Math.Cos(plunge);
                a[2, 1] = 0;
                a[2, 2] = Math.Sin(plunge);

                // Direction cosines of mean pole
                double[] x = new double[3];
                x[0] = Math.Cos(theta) * Math.Cos(phi);
                x[1] = Math.Cos(theta) * Math.Sin(phi);
                x[2] = Math.Sin(theta);

                // Matrix multiplication
                double[] f = new double[3];
                for (int i = 0; i < 3; i++)
                {
                    f[i] = 0;
                    for (int j = 0; j < 3; j++)
                    {
                        f[i] += a[i, j] * x[j];
                    }
                }

                // Calculate appropriate dip angle for random fracture
                double tPlunge, tTrend;

                if (f[2] >= 0)
                {
                    tPlunge = Math.Asin(f[2]);

                    if (f[0] <= 0)
                    {
                        if (f[1] > 0)
                        {
                            // (- +)
                            tTrend = Math.Acos(f[0] / Math.Cos(tPlunge));
                        }
                        else
                        {
                            // (- -)
                            tTrend = Pi - Math.Asin(f[1] / Math.Cos(tPlunge));
                        }
                    }
                    else
                    {
                        if (f[1] <= 0)
                        {
                            // (+ -)
                            tTrend = Math.Asin(f[1] / Math.Cos(tPlunge));
                        }
                        else
                        {
                            // (+ +)
                            tTrend = Math.Acos(f[0] / Math.Cos(tPlunge));
                        }
                    }
                }
                else
                {
                    // If f[2] < 0, flip the direction
                    for (int k = 0; k < 3; k++)
                    {
                        f[k] = -f[k];
                    }

                    tPlunge = Math.Asin(f[2]);

                    if (f[0] <= 0)
                    {
                        if (f[1] > 0)
                        {
                            // (- +)
                            tTrend = Math.Acos(f[0] / Math.Cos(tPlunge)) - Pi;
                        }
                        else
                        {
                            // (- -)
                            tTrend = Math.Asin(f[1] / Math.Cos(tPlunge)) - Pi / 2;
                        }
                    }
                    else
                    {
                        if (f[1] <= 0)
                        {
                            // (+ -)
                            tTrend = Math.Asin(f[1] / Math.Cos(tPlunge));
                        }
                        else
                        {
                            // (+ +)
                            tTrend = Math.Acos(f[0] / Math.Cos(tPlunge)) - Pi;
                        }
                    }
                }

                // Calculate appropriate dip direction for random fracture
                double rndDipDirRad;
                if (tTrend < Pi)
                {
                    rndDipDirRad = tTrend + Pi;
                }
                else
                {
                    rndDipDirRad = tTrend - Pi;
                }

                // Convert radians back to degrees
                rndDip = (Pi / 2 - tPlunge);
                rndDipDir = rndDipDirRad;
            }
            else
            {
                // If Fisher constant K is 0, use mean values directly
                rndDip = meanDip;
                rndDipDir = meanDipDir;
            }
        }

        private double GenerateNormalRandom(double mean, double stdDev, Random rand)
        {
            double u1 = 1.0 - rand.NextDouble();
            double u2 = 1.0 - rand.NextDouble();
            double randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2);
            return mean + stdDev * randStdNormal;
        }

        /// <summary>
        /// Kinematic check - Check kinematic stability based on randomly sampled joint orientations and friction angles
        /// </summary>
        /// <param name="rndFricA">Random friction angle for Set A</param>
        /// <param name="rndDipDirA">Random dip direction for Set A</param>
        /// <param name="rndDipA">Random dip angle for Set A</param>
        /// <param name="rndFricB">Random friction angle for Set B</param>
        /// <param name="rndDipDirB">Random dip direction for Set B</param>
        /// <param name="rndDipB">Random dip angle for Set B</param>
        /// <param name="rndSlopeAngle">Random slope angle</param>
        /// <param name="rndSlopeDipDir">Random slope dip direction</param>
        /// <param name="kineSlide">Output: Kinematic feasibility of sliding</param>
        /// <param name="kineSlideTop">Output: Kinematic feasibility of sliding and toppling</param>
        /// <param name="kineFlex">Output: Kinematic feasibility of flexural toppling</param>
        /// <param name="kineBlockTop">Output: Kinematic feasibility of block toppling</param>
        private void KinematicCheck(
            double rndFricA, double rndDipDirA, double rndDipA,
            double rndFricB, double rndDipDirB, double rndDipB,
            double rndSlopeAngle, double rndSlopeDipDir,
            out bool kineSlide, out bool kineSlideTop,
            out bool kineFlex, out bool kineBlockTop)
        {
            const double Pi = Math.PI;

            // Initialize output variables
            kineSlide = false;
            kineSlideTop = false;
            kineFlex = false;
            kineBlockTop = false;

            // Check if Set A and Set B form a kinematically feasible toppling block
            bool kineTopFeasSetB = false;
            bool kineTopFeasSetA = false;
            bool kineSlideSetA = false;

            // Calculate plunge and trend for Set B
            double plungeB = Pi / 2 - rndDipB;
            double trendB = rndDipDirB + Pi;
            if (trendB > 2 * Pi)
            {
                trendB = trendB - 2 * Pi;
            }

            // Check if friction angle of Set A is greater than its dip angle
            // If not, sliding along Set A takes precedence over block toppling
            if (rndFricA < rndDipA)
            {
                kineSlideSetA = true;
            }

            // Check toppling conditions for Set B
            if (rndSlopeDipDir + (20 * Pi / 180) > 2 * Pi)
            {
                if (trendB < rndSlopeDipDir + (20 * Pi / 180) - 2 * Pi || trendB > rndSlopeDipDir - (20 * Pi / 180))
                {
                    if (plungeB < rndSlopeAngle - rndFricA)
                    {
                        kineTopFeasSetB = true;
                    }
                }
            }
            else if (rndSlopeDipDir - (10 * Pi / 180) < 0)
            {
                if (trendB > rndSlopeDipDir - (20 * Pi / 180) + 2 * Pi || trendB < rndSlopeDipDir + (20 * Pi / 180))
                {
                    if (plungeB < rndSlopeAngle - rndFricA)
                    {
                        kineTopFeasSetB = true;
                    }
                }
            }
            else
            {
                if (trendB > rndSlopeDipDir - (20 * Pi / 180) && trendB < rndSlopeDipDir + (20 * Pi / 180))
                {
                    if (plungeB < rndSlopeAngle - rndFricA)
                    {
                        kineTopFeasSetB = true;
                    }
                }
            }

            // Check necessary toppling conditions for Set A
            if (rndSlopeDipDir + (20 * Pi / 180) > 2 * Pi)
            {
                if (rndDipDirA < rndSlopeDipDir + (20 * Pi / 180) - 2 * Pi || rndDipDirA > rndSlopeDipDir - (20 * Pi / 180))
                {
                    if (rndDipA < rndSlopeAngle)
                    {
                        kineTopFeasSetA = true;
                    }
                }
            }
            else if (rndSlopeDipDir - (20 * Pi / 180) < 0)
            {
                if (rndDipDirA > (rndSlopeDipDir - (20 * Pi / 180) + 2 * Pi) || rndDipDirA < (rndSlopeDipDir + (20 * Pi / 180)))
                {
                    if (rndDipA < rndSlopeAngle)
                    {
                        kineTopFeasSetA = true;
                    }
                }
            }
            else
            {
                if (rndDipDirA > rndSlopeDipDir - (20 * Pi / 180) && rndDipDirA < rndSlopeDipDir + (20 * Pi / 180))
                {
                    if (rndDipA < rndSlopeAngle)
                    {
                        kineTopFeasSetA = true;
                    }
                }
            }

            // Determine which modes are kinematically feasible
            if (kineSlideSetA == false && kineTopFeasSetA == true && kineTopFeasSetB == true)
            {
                // Block toppling conditions are satisfied
                kineBlockTop = true;
                kineSlide = false;
                kineFlex = false;
                kineSlideTop = false;
                _IP.KineBlockTopCount++;
            }
            else if (kineSlideSetA == true && kineTopFeasSetA == true && kineTopFeasSetB == false)
            {
                // Only sliding along Set A is kinematically possible
                kineBlockTop = false;
                kineSlide = true;
                kineFlex = false;
                kineSlideTop = false;
                _IP.KineSlideCount++;
            }
            else if (kineTopFeasSetA == false && kineTopFeasSetB == true)
            {
                // Flexural/block flexural toppling is kinematically possible
                kineBlockTop = false;
                kineSlide = false;
                kineFlex = true;
                kineSlideTop = false;
                _IP.KineFlexCount++;
            }
            else if (kineSlideSetA == true && kineTopFeasSetA == true && kineTopFeasSetB == true)
            {
                // Sliding and toppling will occur
                kineBlockTop = false;
                kineSlide = false;
                kineFlex = false;
                kineSlideTop = true;
                _IP.KineSlideTopCount++;
            }
            else
            {
                // No kinematic failure possible
                kineBlockTop = false;
                kineSlide = false;
                kineFlex = false;
                kineSlideTop = false;
            }
        }
        private void LimitEquilibrium(
                  double rndSlopeAngle, double rndTopAngle, double rndHeight, double rndSlopeDipDir,
                  double rndDipA, double rndDipB, double rndFricA, double rndFricB,
                  double rndUnitWt, double rndUnitWtH2O, double rndSeis, double rndPorePress,
                  ref double rndSpaceA, ref double rndSpaceB,
                  ref double lambdaSpaceA, ref double zetaSpaceA,
                  ref double lambdaSpaceB, ref double zetaSpaceB,
                  ref bool printIteration, int ini)
        {
            Random rand = new Random();

            // Define temporary variables for block generation
            double x1, y1, x2, y2, x3, y3, x4, y4;
            double xUS, yUS, xDS, yDS;
            int blockCount = 0;
            bool crestBlock;

            // Create lists to store calculation values
            List<BlockXY> blockCoords = new List<BlockXY>();
            List<PorePressXY> waterLevel = new List<PorePressXY>();
            List<double> blockArea = new List<double>();        // Area of each block
            List<double> centreOfMassX = new List<double>();    // X coordinate of center of mass
            List<double> centreOfMassY = new List<double>();    // Y coordinate of center of mass
            List<double> blockWeight = new List<double>();      // Weight of each block
            List<double> blockWeightArm = new List<double>();   // Moment arm of block weight
            List<double> v1 = new List<double>();              // Downslope thrust force
            List<double> v1arm = new List<double>();           // Moment arm of v1
            List<double> v2 = new List<double>();              // Pore water pressure
            List<double> v2arm = new List<double>();           // Moment arm of v2
            List<double> v3 = new List<double>();              // Upslope thrust force
            List<double> v3arm = new List<double>();           // Moment arm of v3
            List<double> mn = new List<double>();              // Moment arm of Pn
            List<double> ln = new List<double>();              // Moment arm of Pn-1
            List<double> pnTanArm = new List<double>();        // Moment arm of PnTan
            List<double> fdSeismic = new List<double>();       // Seismic force
            List<double> seismicArm = new List<double>();      // Moment arm of seismic force

            // Establish overall slope geometry
            double toeX = 0;
            double toeY = 0;
            double crestX = rndHeight / Math.Tan(rndSlopeAngle);
            double crestY = rndHeight;
            double rightLimX = Math.Max(crestX + rndHeight / Math.Tan(rndDipA), _IP.MeanSpaceA * Math.Cos(_IP.MeanDipA));
            double rightLimY = crestY + (rightLimX - crestX) * Math.Tan(rndTopAngle);
            double leftLimX = -10;
            double leftLimY = toeY;

            // Define starting point at slope toe
            double botLeftX = toeX;
            double botLeftY = toeY;
            double topLeftX = toeX;
            double topLeftY = toeY;

            // Define pore water conditions at toe
            double pwLeftX = 0;
            double pwLeftY = 0;


            ExcelWorksheet worksheet3 = package.Workbook.Worksheets["Analysis Details 3"];
            ExcelWorksheet worksheet2 = package.Workbook.Worksheets["Analysis Details 2"];
            // Iteratively generate blocks
            do
            {
                // Generate random spacing for set A based on distribution type
                if (_IP.DistSpaceA == "Normal")
                {
                    do
                    {
                        rndSpaceA = GenerateNormalRandom(_IP.MeanSpaceA, _IP.StDevSpaceA, rand);
                    } while (rndSpaceA <= 0);
                }
                else if (_IP.DistSpaceA == "Lognormal")
                {
                    zetaSpaceA = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevSpaceA, 2) / Math.Pow(_IP.MeanSpaceA, 2))));
                    lambdaSpaceA = Math.Log(_IP.MeanSpaceA) - 0.5 * Math.Pow(zetaSpaceA, 2);
                    rndSpaceA = Math.Exp(GenerateNormalRandom(lambdaSpaceA, zetaSpaceA, rand));
                }
                else if (_IP.DistSpaceA == "Exponential")
                {
                    rndSpaceA = -_IP.MeanSpaceA * Math.Log(1 - rand.NextDouble());
                }
                else if (_IP.DistSpaceA == "Fixed")
                {
                    rndSpaceA = _IP.MeanSpaceA;
                }

                // Generate random spacing for set B based on distribution type
                if (_IP.DistSpaceB == "Normal")
                {
                    do
                    {
                        rndSpaceB = GenerateNormalRandom(_IP.MeanSpaceB, _IP.StDevSpaceB, rand);
                    } while (rndSpaceB <= 0);
                }
                else if (_IP.DistSpaceB == "Lognormal")
                {
                    zetaSpaceB = Math.Sqrt(Math.Log(1 + (Math.Pow(_IP.StDevSpaceB, 2) / Math.Pow(_IP.MeanSpaceB, 2))));
                    lambdaSpaceB = Math.Log(_IP.MeanSpaceB) - 0.5 * Math.Pow(zetaSpaceB, 2);
                    rndSpaceB = Math.Exp(GenerateNormalRandom(lambdaSpaceB, zetaSpaceB, rand));
                }
                else if (_IP.DistSpaceB == "Exponential")
                {
                    rndSpaceB = -_IP.MeanSpaceB * Math.Log(1 - rand.NextDouble());
                }
                else if (_IP.DistSpaceB == "Fixed")
                {
                    rndSpaceB = _IP.MeanSpaceB;
                }

                //Widen Set B Spacing for blocks below the slope crest if block bolting is specified
                if (_IP.BoltBlocksTogether == true && topLeftY < crestY)
                    rndSpaceB = rndSpaceB * _IP.EffectiveWidth;

                // Calculate coordinates for the right side of current block
                double botRightY = botLeftY + (rndSpaceB / Math.Cos(rndDipA - (Math.PI / 2 - rndDipB))) * Math.Sin(rndDipA);
                double botRightX = botLeftX + (rndSpaceB / Math.Cos(rndDipA - (Math.PI / 2 - rndDipB))) * Math.Cos(rndDipA);
                double topRightY = topLeftY + (rndSpaceB / Math.Cos(rndSlopeAngle - (Math.PI / 2 - rndDipB))) * Math.Sin(rndSlopeAngle);
                double topRightX = topLeftX + (rndSpaceB / Math.Cos(rndSlopeAngle - (Math.PI / 2 - rndDipB))) * Math.Cos(rndSlopeAngle);

                double pwRightX, pwRightY;
                double hypDist, xDist;

                // Determine position of current block relative to crest
                if (topRightY < crestY && topRightX < crestX)
                {
                    // Block is below crest
                    crestBlock = false;
                    pwRightX = botRightX - rndPorePress * (botRightX - topRightX);
                    pwRightY = botRightY + rndPorePress * (topRightY - botRightY);
                }
                else if (topLeftX < crestX && topRightX > crestX)
                {
                    // Block crosses crest
                    crestBlock = true;
                    hypDist = Math.Sqrt(Math.Pow(crestX - topLeftX, 2) + Math.Pow(crestY - topLeftY, 2));
                    xDist = hypDist * Math.Cos(rndSlopeAngle - (Math.PI / 2 - rndDipB));

                    topRightX = crestX + ((rndSpaceB - xDist) / Math.Cos(Math.PI / 2 - rndDipB - rndTopAngle)) * Math.Cos(rndTopAngle);
                    topRightY = crestY + ((rndSpaceB - xDist) / Math.Cos(Math.PI / 2 - rndDipB - rndTopAngle)) * Math.Sin(rndTopAngle);

                    pwRightX = botRightX - rndPorePress * (botRightX - topRightX);
                    pwRightY = botRightY + rndPorePress * (topRightY - botRightY);
                }
                else
                {
                    // Block is above crest
                    crestBlock = false;
                    topRightX = topLeftX + (rndSpaceB / Math.Cos(rndTopAngle - (Math.PI / 2 - rndDipB))) * Math.Cos(rndTopAngle);
                    topRightY = topLeftY + (rndSpaceB / Math.Cos(rndTopAngle - (Math.PI / 2 - rndDipB))) * Math.Sin(rndTopAngle);

                    pwRightX = botRightX - rndPorePress * (botRightX - topRightX);
                    pwRightY = botRightY + rndPorePress * (topRightY - botRightY);
                }

                // Conditions for ending block formation at slope top
                if (botRightY > topRightY || botLeftY > topLeftY)
                {
                    break;
                }

                // Add calculated coordinates to lists
                blockCoords.Add(new BlockXY
                {
                    xTL = topLeftX,
                    yTL = topLeftY,
                    xBL = botLeftX,
                    yBL = botLeftY,
                    xBR = botRightX,
                    yBR = botRightY,
                    xTR = topRightX,
                    yTR = topRightY
                });

                waterLevel.Add(new PorePressXY
                {
                    xDS = pwLeftX,
                    yDS = pwLeftY,
                    xUS = pwRightX,
                    yUS = pwRightY
                });

                // Calculate forces and corresponding moment arms
                x1 = blockCoords[blockCount].xTL;
                y1 = blockCoords[blockCount].yTL;
                x2 = blockCoords[blockCount].xBL;
                y2 = blockCoords[blockCount].yBL;
                x3 = blockCoords[blockCount].xBR;
                y3 = blockCoords[blockCount].yBR;
                x4 = blockCoords[blockCount].xTR;
                y4 = blockCoords[blockCount].yTR;
                xUS = waterLevel[blockCount].xUS;
                yUS = waterLevel[blockCount].yUS;
                xDS = waterLevel[blockCount].xDS;
                yDS = waterLevel[blockCount].yDS;

                // Calculate block area and center of mass
                double area;
                if (crestBlock)
                {
                    area = 0.5 * ((x1 * y2 - x2 * y1) + (x2 * y3 - x3 * y2) + (x3 * y4 - x4 * y3) +
                                  (x4 * crestY - crestX * y4) + (crestX * y1 - x1 * crestY));
                    double comY = (1 / (6 * area)) * ((y1 + y2) * (x1 * y2 - x2 * y1) +
                        (y2 + y3) * (x2 * y3 - x3 * y2) + (y3 + y4) * (x3 * y4 - x4 * y3) +
                        (y4 + crestY) * (x4 * crestY - crestX * y4) + (crestY + y1) * (crestX * y1 - x1 * crestY));
                    double comX = (1 / (6 * area)) * ((x1 + x2) * (x1 * y2 - x2 * y1) +
                        (x2 + x3) * (x2 * y3 - x3 * y2) + (x3 + x4) * (x3 * y4 - x4 * y3) +
                        (x4 + crestX) * (x4 * crestY - crestX * y4) + (crestX + x1) * (crestX * y1 - x1 * crestY));
                    centreOfMassX.Add(comX);
                    centreOfMassY.Add(comY);
                }
                else
                {
                    area = 0.5 * ((x1 * y2 - y1 * x2) + (x2 * y3 - x3 * y2) +
                                  (x3 * y4 - x4 * y3) + (x4 * y1 - x1 * y4));
                    double comY = (1 / (6 * area)) * ((y1 + y2) * (x1 * y2 - x2 * y1) +
                        (y2 + y3) * (x2 * y3 - x3 * y2) + (y3 + y4) * (x3 * y4 - x4 * y3) +
                        (y4 + y1) * (x4 * y1 - x1 * y4));
                    double comX = (1 / (6 * area)) * ((x1 + x2) * (x1 * y2 - x2 * y1) +
                        (x2 + x3) * (x2 * y3 - x3 * y2) + (x3 + x4) * (x3 * y4 - x4 * y3) +
                        (x4 + x1) * (x4 * y1 - x1 * y4));
                    centreOfMassX.Add(comX);
                    centreOfMassY.Add(comY);
                }
                blockArea.Add(area);

                // Calculate block weight and moment arm
                blockWeight.Add(area * rndUnitWt);
                blockWeightArm.Add(centreOfMassX[blockCount] - x2);
                mn.Add(Math.Sqrt(Math.Pow(x3 - x4, 2) + Math.Pow(y3 - y4, 2)) +
                       rndSpaceB * Math.Tan(rndDipA - (Math.PI / 2 - rndDipB)));
                ln.Add(Math.Sqrt(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2)));
                pnTanArm.Add(rndSpaceB);
                fdSeismic.Add(blockWeight[blockCount] * rndSeis);
                seismicArm.Add(centreOfMassY[blockCount] - y2);

                // Calculate water pressure
                if (yUS > y3)
                {
                    v1.Add(0.5 * Math.Pow(yUS - y3, 2) * rndUnitWtH2O);
                    v1arm.Add((1.0 / 3.0) * Math.Sqrt(Math.Pow(x3 - xUS, 2) + Math.Pow(y3 - yUS, 2)) +
                             rndSpaceB * Math.Tan(rndDipA - (Math.PI / 2 - rndDipB)));
                }
                else
                {
                    v1.Add(0);
                    v1arm.Add(0);
                }

                if (yDS > y2)
                {
                    v3.Add(0.5 * Math.Pow(yDS - y2, 2) * rndUnitWtH2O);
                    v3arm.Add((1.0 / 3.0) * Math.Sqrt(Math.Pow(xDS - x2, 2) + Math.Pow(yDS - y2, 2)));
                }
                else
                {
                    v3.Add(0);
                    v3arm.Add(0);
                }

                // Calculate base water pressure
                if (yDS < y2 && yUS < y3)
                {
                    v2.Add(0);
                    v2arm.Add(0);
                }
                else if (yDS >= y2 && yUS > y3)
                {
                    double baseLength = Math.Sqrt(Math.Pow(x3 - x2, 2) + Math.Pow(y3 - y2, 2));
                    v2.Add((((yUS - y3) - (yDS - y2)) / 2 + (yDS - y2)) * baseLength * rndUnitWtH2O);
                    if (yUS - y3 > yDS - y2)
                    {
                        v2arm.Add(rndUnitWtH2O * Math.Pow(baseLength, 2) *
                                 ((yDS - y2) / 2 + ((yUS - y3) - (yDS - y2)) / 3) / v2[blockCount]);
                    }
                    else
                    {
                        v2arm.Add(rndUnitWtH2O * Math.Pow(baseLength, 2) *
                                 ((yDS - y2) / 2 + ((yUS - y3) - (yDS - y2)) / 6) / v2[blockCount]);
                    }
                }
                else if (yDS < y2 && yUS > y3)
                {
                    double theta = Math.Atan(Math.Abs(yUS - yDS) / Math.Abs(xUS - xDS)) - rndDipA;
                    double alpha = Math.PI - theta - (rndDipA + rndDipB);
                    double sideLength = Math.Sqrt(Math.Pow(yUS - y3, 2) + Math.Pow(xUS - x3, 2));
                    double lengthAct = (Math.Sin(alpha) / Math.Sin(theta)) * sideLength;
                    v2.Add(0.5 * (yUS - y3) * rndUnitWtH2O * lengthAct);
                    v2arm.Add(Math.Sqrt(Math.Pow(x3 - x2, 2) + Math.Pow(y3 - y2, 2)) - (1.0 / 3.0) * lengthAct);
                }
                else if (yDS > y2 && yUS < y3)
                {
                    double theta = Math.Atan(Math.Abs(yDS - yUS) / Math.Abs(xUS - xDS)) + rndDipA;
                    double alpha = Math.PI - (Math.PI - rndDipB - rndDipA) - theta;
                    double sideLength = Math.Sqrt(Math.Pow(yDS - y2, 2) + Math.Pow(xDS - x2, 2));
                    double lengthAct = (Math.Sin(alpha) / Math.Sin(theta)) * sideLength;
                    v2.Add(0.5 * (yDS - y2) * rndUnitWtH2O * lengthAct);
                    v2arm.Add((1.0 / 3.0) * lengthAct);
                }

                // Update left coordinates for next block
                botLeftX = blockCoords[blockCount].xBR - (rndSpaceA / Math.Sin(Math.PI - rndDipA - rndDipB) *
                          Math.Sin(Math.PI / 2 - rndDipB));
                botLeftY = blockCoords[blockCount].yBR + (rndSpaceA / Math.Sin(Math.PI - rndDipA - rndDipB) *
                          Math.Cos(Math.PI / 2 - rndDipB));
                topLeftX = blockCoords[blockCount].xTR;
                topLeftY = blockCoords[blockCount].yTR;
                pwLeftX = waterLevel[blockCount].xUS;
                pwLeftY = waterLevel[blockCount].yUS;

                blockCount++;
            } while (true);

            // Write calculated slope geometry parameters to "Analysis Details 3" worksheet
            if (printIteration)
            {
                try
                {
                    const double Pi = Math.PI;

                    // Print iteration number
                    worksheet3.Cells[1, 3].Value = ini;

                    // Print values of randomly sampled parameters
                    worksheet3.Cells[2, 2].Value = blockCount;
                    worksheet3.Cells[3, 2].Value = rndHeight;
                    worksheet3.Cells[4, 2].Value = rndSlopeAngle * (180.0 / Pi);
                    worksheet3.Cells[5, 2].Value = rndTopAngle * (180.0 / Pi);
                    worksheet3.Cells[6, 2].Value = rndDipA * (180.0 / Pi);
                    worksheet3.Cells[7, 2].Value = rndDipB * (180.0 / Pi);
                    worksheet3.Cells[8, 2].Value = rndUnitWt;
                    worksheet3.Cells[9, 2].Value = rndUnitWtH2O;
                    worksheet3.Cells[10, 2].Value = rndSeis;
                    worksheet3.Cells[11, 2].Value = rndPorePress;

                    // Print coordinates defining overall slope geometry
                    worksheet3.Cells[17, 2].Value = leftLimX;
                    worksheet3.Cells[18, 2].Value = toeX;
                    worksheet3.Cells[19, 2].Value = crestX;
                    worksheet3.Cells[20, 2].Value = rightLimX;
                    worksheet3.Cells[17, 3].Value = leftLimY;
                    worksheet3.Cells[18, 3].Value = toeY;
                    worksheet3.Cells[19, 3].Value = crestY;
                    worksheet3.Cells[20, 3].Value = rightLimY;

                    if (_IP.AddToeSupport)
                    {
                        _IP.SupportForce = _IP.Magnitude;
                        worksheet3.Cells[12, 2].Value = _IP.SupportForce;
                    }
                    else
                    {
                        _IP.SupportForce = 0;
                        worksheet3.Cells[12, 2].Value = _IP.SupportForce;
                    }

                    if (_IP.BoltBlocksTogether)
                    {
                        worksheet3.Cells[13, 2].Value = "Yes";
                    }
                    else
                    {
                        worksheet3.Cells[13, 2].Value = "No";
                    }

                    // Print block coordinates and other calculation parameters
                    int inrow2 = 1;
                    for (int j = 0; j < blockCount; j++)
                    {
                        // Print block coordinates
                        worksheet3.Cells[inrow2 + 22, 1].Value = $"Block {j + 1}";
                        worksheet3.Cells[inrow2 + 22, 2].Value = blockCoords[j].xTL;
                        worksheet3.Cells[inrow2 + 22, 3].Value = blockCoords[j].yTL;
                        worksheet3.Cells[inrow2 + 23, 2].Value = blockCoords[j].xBL;
                        worksheet3.Cells[inrow2 + 23, 3].Value = blockCoords[j].yBL;
                        worksheet3.Cells[inrow2 + 24, 2].Value = blockCoords[j].xBR;
                        worksheet3.Cells[inrow2 + 24, 3].Value = blockCoords[j].yBR;
                        worksheet3.Cells[inrow2 + 25, 2].Value = blockCoords[j].xTR;
                        worksheet3.Cells[inrow2 + 25, 3].Value = blockCoords[j].yTR;

                        // Print other calculation parameters
                        worksheet3.Cells[inrow2 + 22, 6].Value = centreOfMassX[j];
                        worksheet3.Cells[inrow2 + 22, 7].Value = centreOfMassY[j];
                        worksheet3.Cells[inrow2 + 22, 8].Value = blockWeight[j];
                        worksheet3.Cells[inrow2 + 22, 9].Value = fdSeismic[j];
                        worksheet3.Cells[inrow2 + 22, 10].Value = v1[j];
                        worksheet3.Cells[inrow2 + 22, 11].Value = v2[j];
                        worksheet3.Cells[inrow2 + 22, 12].Value = v3[j];
                        worksheet3.Cells[inrow2 + 22, 13].Value = v1arm[j];
                        worksheet3.Cells[inrow2 + 22, 14].Value = v2arm[j];
                        worksheet3.Cells[inrow2 + 22, 15].Value = v3arm[j];
                        worksheet3.Cells[inrow2 + 22, 16].Value = seismicArm[j];
                        worksheet3.Cells[inrow2 + 22, 17].Value = blockWeightArm[j];
                        worksheet3.Cells[inrow2 + 22, 18].Value = pnTanArm[j];
                        worksheet3.Cells[inrow2 + 22, 19].Value = mn[j];
                        worksheet3.Cells[inrow2 + 22, 20].Value = ln[j];
                        worksheet3.Cells[inrow2 + 22, 21].Value = rndFricA * (180.0 / Pi);
                        worksheet3.Cells[inrow2 + 22, 22].Value = rndFricB * (180.0 / Pi);

                        inrow2 += 4;
                    }

                    // Print water table coordinates
                    inrow2 = 1;
                    if (rndPorePress != 0) // Don't display water table if there's no pore pressure
                    {
                        for (int j = 0; j < blockCount; j++)
                        {
                            worksheet3.Cells[inrow2 + 22, 4].Value = waterLevel[j].xDS;
                            worksheet3.Cells[inrow2 + 22, 5].Value = waterLevel[j].yDS;
                            worksheet3.Cells[inrow2 + 23, 4].Value = waterLevel[j].xUS;
                            worksheet3.Cells[inrow2 + 23, 5].Value = waterLevel[j].yUS;

                            inrow2 += 4;
                        }

                        worksheet3.Cells[inrow2 + 20, 4].Value = rightLimX;
                        worksheet3.Cells[inrow2 + 20, 5].Value = rightLimY;
                    }
                    else
                    {
                        worksheet3.Cells[inrow2 + 22, 4].Value = "N/A";
                    }

                    // Scale preview image
                    double maxScale;
                    if (blockCount < 2)
                    {
                        maxScale = Math.Ceiling(rightLimY / 10.0) * 10.0;
                    }
                    else
                    {
                        maxScale = Math.Max(
                            Math.Ceiling((blockCoords[blockCount - 1].xBR + 5) / 10.0) * 10.0,
                            Math.Ceiling((blockCoords[blockCount - 1].yTR + 5) / 10.0) * 10.0
                        );
                    }

                    if (worksheet3.Drawings.Count > 0)
                    {
                        var chart = worksheet3.Drawings[0] as ExcelChart;
                        if (chart != null)
                        {
                            // Set horizontal axis (X-axis) range
                            chart.XAxis.MinValue = leftLimX;
                            chart.XAxis.MaxValue = maxScale;
                            chart.XAxis.MajorUnit = 10;

                            // Set vertical axis (Y-axis) range
                            chart.YAxis.MinValue = leftLimY;
                            chart.YAxis.MaxValue = maxScale;
                            chart.YAxis.MajorUnit = 10;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error writing slope geometry parameters to \"Analysis Details 3\" Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // *******************************************************************************************
            // Calculate block interaction forces
            // *******************************************************************************************
            // Interaction force variables
            double rn;                   // Normal force at block base
            double sn;                   // Shear force at block base
            double a1;                   // Irregular geometry adjustment angle
            bool sliding = false;        // Sliding flag

            List<double> pnMin1Topple = new List<double>();  // Pn-1 for toppling
            List<double> pnMin1Slide = new List<double>();   // Pn-1 for sliding
            List<double> forcePnMin1 = new List<double>();   // Pn-1 force (minimum of toppling and sliding)
            List<double> pn = new List<double>();           // Pn force
            List<double> rnArray = new List<double>();       // Array of normal forces
            List<double> snArray = new List<double>();       // Array of shear forces
            List<string> failureMode = new List<string>();  // Block failure mode (stable, sliding, or toppling)

            // Initialize lists, each list only needs one initial element
            pnMin1Topple.Add(0);
            pnMin1Slide.Add(0);
            forcePnMin1.Add(0);
            pn.Add(0);
            rnArray.Add(0);
            snArray.Add(0);
            failureMode.Add("");

            // Pn for the uppermost block equals 0
            pn[pn.Count - 1] = 0;

            // Angle adjustment for irregular geometry
            a1 = (rndDipA - (Math.PI / 2 - rndDipB));

            // Calculate block interaction forces from top to bottom
            for (int i = blockCount - 1; i >= 0; i--)
            {
                if (ln[i] == 0) // Reached the toe block
                {
                    pnMin1Slide[pnMin1Slide.Count - 1] = 0;
                    pnMin1Topple[pnMin1Topple.Count - 1] = 0;
                    forcePnMin1[forcePnMin1.Count - 1] = 0;
                    failureMode[failureMode.Count - 1] = "See Analysis Details 2";
                }
                else
                {
                    // Calculate forces needed to prevent block i from toppling and sliding
                    pnMin1Topple[pnMin1Topple.Count - 1] = ((-blockWeight[i] * blockWeightArm[i]) + fdSeismic[i] * seismicArm[i] -
                        pn[pn.Count - 1] * Math.Tan(rndFricB) * pnTanArm[i] - v3[i] * v3arm[i] +
                        v1[i] * v1arm[i] + v2[i] * v2arm[i] + pn[pn.Count - 1] * mn[i]) / ln[i];

                    pnMin1Slide[pnMin1Slide.Count - 1] = pn[pn.Count - 1] - (blockWeight[i] * (Math.Cos(rndDipA) *
                        Math.Tan(rndFricA) - Math.Sin(rndDipA)) - fdSeismic[i] * (Math.Sin(rndDipA) * Math.Tan(rndFricA) +
                        Math.Cos(rndDipA)) - v2[i] * Math.Tan(rndFricA) + (v3[i] - v1[i]) * (Math.Cos(a1) + Math.Sin(a1) *
                        Math.Tan(rndFricA))) / (Math.Sin(a1) * (Math.Tan(rndFricA) + Math.Tan(rndFricB)) + Math.Cos(a1) *
                        (1 - Math.Tan(rndFricA) * Math.Tan(rndFricB)));

                    if (pnMin1Slide[pnMin1Slide.Count - 1] < 0 && pnMin1Topple[pnMin1Topple.Count - 1] < 0)
                    {
                        // Block is stable
                        forcePnMin1[forcePnMin1.Count - 1] = 0;
                        failureMode[failureMode.Count - 1] = "Stable";
                        sliding = false;
                    }
                    else if (pnMin1Topple[pnMin1Topple.Count - 1] > pnMin1Slide[pnMin1Slide.Count - 1] && !sliding)
                    {
                        // Toppling is critical mode, check if base is sliding
                        rn = blockWeight[i] * Math.Cos(rndDipA) - fdSeismic[i] * Math.Sin(rndDipA) - v2[i] +
                            (v3[i] - v1[i]) * Math.Sin(a1) + (pnMin1Topple[pnMin1Topple.Count - 1] - pn[pn.Count - 1]) * Math.Sin(a1) +
                            (pn[pn.Count - 1] - pnMin1Topple[pnMin1Topple.Count - 1]) * (Math.Tan(rndFricB) * Math.Cos(a1));

                        sn = blockWeight[i] * Math.Sin(rndDipA) + fdSeismic[i] * Math.Cos(rndDipA) +
                            (v1[i] - v3[i]) * Math.Cos(a1) + (pn[pn.Count - 1] - pnMin1Topple[pnMin1Topple.Count - 1]) * Math.Cos(a1) +
                            (pn[pn.Count - 1] - pnMin1Topple[pnMin1Topple.Count - 1]) * (Math.Tan(rndFricB) * Math.Sin(a1));

                        if (rn > 0 && Math.Abs(sn) < rn * Math.Tan(rndFricA))
                        {
                            // Base is not sliding
                            forcePnMin1[forcePnMin1.Count - 1] = pnMin1Topple[pnMin1Topple.Count - 1];
                            failureMode[failureMode.Count - 1] = "Toppling";
                        }
                        else
                        {
                            // Base is sliding
                            if (pnMin1Slide[pnMin1Slide.Count - 1] < 0)
                            {
                                forcePnMin1[forcePnMin1.Count - 1] = 0;
                            }
                            else
                            {
                                forcePnMin1[forcePnMin1.Count - 1] = pnMin1Slide[pnMin1Slide.Count - 1];
                            }
                            failureMode[failureMode.Count - 1] = "Sliding";
                            sliding = true;
                        }
                    }
                    else
                    {
                        // Sliding is critical mode
                        if (pnMin1Slide[pnMin1Slide.Count - 1] < 0)
                        {
                            forcePnMin1[forcePnMin1.Count - 1] = 0;
                        }
                        else
                        {
                            forcePnMin1[forcePnMin1.Count - 1] = pnMin1Slide[pnMin1Slide.Count - 1];
                        }
                        failureMode[failureMode.Count - 1] = "Sliding";
                        sliding = true;
                    }
                }

                // Calculate base normal force and shear force
                rnArray[rnArray.Count - 1] = blockWeight[i] * Math.Cos(rndDipA) - fdSeismic[i] * Math.Sin(rndDipA) - v2[i] +
                    (v3[i] - v1[i]) * Math.Sin(a1) + (forcePnMin1[forcePnMin1.Count - 1] - pn[pn.Count - 1]) * Math.Sin(a1) +
                    (pn[pn.Count - 1] - forcePnMin1[forcePnMin1.Count - 1]) * (Math.Tan(rndFricB) * Math.Cos(a1));

                snArray[snArray.Count - 1] = blockWeight[i] * Math.Sin(rndDipA) + fdSeismic[i] * Math.Cos(rndDipA) +
                    (v1[i] - v3[i] + pn[pn.Count - 1] - forcePnMin1[forcePnMin1.Count - 1]) * Math.Cos(a1) +
                    (pn[pn.Count - 1] - forcePnMin1[forcePnMin1.Count - 1]) * (Math.Tan(rndFricB) * Math.Sin(a1));

                // Print results to Analysis Details 3 worksheet
                if (printIteration)
                {
                    try
                    {
                        if (worksheet3 != null)
                        {
                            worksheet3.Cells[23 + i * 4, 23].Value = pn[pn.Count - 1];
                            worksheet3.Cells[23 + i * 4, 24].Value = pnMin1Topple[pnMin1Topple.Count - 1];
                            worksheet3.Cells[23 + i * 4, 25].Value = pnMin1Slide[pnMin1Slide.Count - 1];
                            worksheet3.Cells[23 + i * 4, 26].Value = forcePnMin1[forcePnMin1.Count - 1];
                            worksheet3.Cells[23 + i * 4, 27].Value = rnArray[rnArray.Count - 1];
                            worksheet3.Cells[23 + i * 4, 28].Value = snArray[snArray.Count - 1];
                            worksheet3.Cells[23 + i * 4, 29].Value = sliding;
                            worksheet3.Cells[23 + i * 4, 30].Value = failureMode[failureMode.Count - 1];
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error writing block interaction forces to Analysis Details 3: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Extend all lists
                failureMode.Add("");
                pnMin1Topple.Add(0);
                pnMin1Slide.Add(0);
                forcePnMin1.Add(0);
                pn.Add(0);
                rnArray.Add(0);
                snArray.Add(0);

                // Set Pn to ForcePnMin1 as in VB code
                pn[pn.Count - 1] = forcePnMin1[forcePnMin1.Count - 2];
            }

            // *******************************************************************************************
            // Toe block analysis (or all stable toe blocks)
            // *******************************************************************************************
            // Toe block analysis variables
            string toeMode = "";
            object toeFS;
            double drivingF = 0, resistingF = 0;
            double drivingM = 0, resistingM = 0;
            double totSn = 0, totRn = 0;
            object fofSSliding, fofSToppling;
            double supportMomentArm = 0;
            double supportMoment = 0;
            int numStableBlock = 0;

            double supportForce = 0;
            double inclination = 0;

            try
            {
                if (_IP.AddToeSupport)
                {
                    supportForce = _IP.Magnitude;
                    inclination = _IP.Orientation;
                    supportMomentArm = 0.5 * Math.Sqrt(Math.Pow(blockCoords[0].xTL - blockCoords[0].xTR, 2) +
                                                      Math.Pow(blockCoords[0].yTL - blockCoords[0].yTR, 2));
                    supportMoment = supportForce * Math.Cos((Math.PI / 2) - rndSlopeAngle - inclination) * supportMomentArm;
                }
                else
                {
                    supportForce = 0;
                    inclination = 0;
                    supportMoment = 0;
                }

                // Calculate number of stable blocks behind toe block and sum Rn and Sn values
                if (blockCount <= 1)
                {
                    numStableBlock = 0;
                }
                else
                {
                    totSn = 0;
                    totRn = 0;
                    for (int i = blockCount - 2; i >= 0; i--) // Start from toe block and move upward
                    {
                        if (failureMode[i] == "Stable")
                        {
                            numStableBlock++;
                            totSn += snArray[i];
                            totRn += rnArray[i];
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                // Calculate driving/resisting forces and moments
                string supportType = _IP.NatureOfForceApplication; // Default active support

                if (supportType == "Active") // Active support
                {
                    // Use index logic consistent with VB
                    drivingF = snArray[snArray.Count - 2] - supportForce * Math.Cos(rndDipA + inclination);
                    resistingF = (rnArray[rnArray.Count - 2] + supportForce * Math.Sin(rndDipA + inclination)) * Math.Tan(rndFricA);
                    drivingM = pn[pn.Count - 2] * mn[0] + v1[0] * v1arm[0] + v2[0] * v2arm[0] +
                               fdSeismic[0] * seismicArm[0] - supportMoment;
                    resistingM = pn[pn.Count - 2] * Math.Tan(rndFricB) * pnTanArm[0] +
                                 blockWeight[0] * blockWeightArm[0];
                }
                else if (supportType == "Passive") // Passive support
                {
                    // Note: VB version uses UBound(RnArray) - 1 as index here
                    drivingF = snArray[snArray.Count - 2];
                    resistingF = (rnArray[rnArray.Count - 2] + supportForce * Math.Sin(rndDipA + inclination)) *
                                 Math.Tan(rndFricA) + supportForce * Math.Cos(rndDipA + inclination);
                    drivingM = pn[pn.Count - 2] * mn[0] + v1[0] * v1arm[0] + v2[0] * v2arm[0] +
                               fdSeismic[0] * seismicArm[0];
                    resistingM = pn[pn.Count - 2] * Math.Tan(rndFricB) * pnTanArm[0] +
                                 blockWeight[0] * blockWeightArm[0] + supportMoment;
                }

                // Determine slope stability based on number of stable blocks
                if (numStableBlock == 0)
                {
                    // Slope stability controlled by toe block, possible toppling or sliding
                    if (drivingF > 0 && resistingF > 0)
                    {
                        fofSSliding = resistingF / drivingF;
                    }
                    else if (drivingF > 0 && resistingF < 0)
                    {
                        fofSSliding = 0;
                    }
                    else
                    {
                        fofSSliding = "Inf.";
                    }

                    if (drivingM > 0)
                    {
                        fofSToppling = resistingM / drivingM;
                    }
                    else
                    {
                        fofSToppling = "Inf.";
                    }

                    // Determine the most critical failure mode and corresponding safety factor for toe block
                    if (fofSToppling.ToString() == "Inf." && fofSSliding.ToString() == "Inf.")
                    {
                        // Both safety factors tend to infinity, failure is impossible
                        toeMode = "No Failure Possible";
                        toeFS = "-";
                    }
                    else if (fofSSliding.ToString() == "Inf." ||
                            (fofSToppling is double && fofSSliding is double &&
                             (double)fofSToppling < (double)fofSSliding && (double)fofSSliding > 1 && blockCount < 2))
                    {
                        // Only one block in slope, and toppling mode is more critical
                        toeFS = fofSToppling;
                        toeMode = "Toppling";
                    }
                    else if (fofSToppling.ToString() == "Inf." ||
                            (fofSToppling is double && fofSSliding is double &&
                             (double)fofSToppling > (double)fofSSliding && blockCount < 2))
                    {
                        // Only one block in slope, and sliding mode is more critical
                        toeFS = fofSSliding;
                        toeMode = "Sliding";
                    }
                    else if (fofSSliding.ToString() == "Inf." ||
                            (fofSToppling is double && fofSSliding is double &&
                             (double)fofSToppling < (double)fofSSliding && (double)fofSSliding > 1 &&
                             failureMode[failureMode.Count - 2] == "Toppling"))
                    {
                        toeFS = fofSToppling;
                        toeMode = "Toppling";
                    }
                    else if (fofSSliding.ToString() == "Inf." ||
                            (fofSToppling is double && fofSSliding is double &&
                             (double)fofSToppling < (double)fofSSliding && (double)fofSSliding > 1 &&
                             failureMode[failureMode.Count - 2] == "Stable"))
                    {
                        toeFS = fofSToppling;
                        toeMode = "Toppling";
                    }
                    else
                    {
                        // Sliding mode is more critical
                        toeMode = "Sliding";
                        toeFS = fofSSliding;
                    }
                }
                else
                {
                    // Stability controlled by stable block group at toe
                    // Sum Rn and Sn for all stable blocks above toe block
                    resistingF = resistingF + totRn * Math.Tan(rndFricA);
                    drivingF = drivingF + totSn;
                    fofSSliding = resistingF / drivingF;
                    fofSToppling = "-";
                    toeFS = fofSSliding;
                    toeMode = "Sliding";
                }

                // Count stable/unstable iterations
                if (toeMode == "Toppling")
                {
                    _IP.ToppleCount++;
                    if (fofSToppling is double && (double)fofSToppling > 1)
                    {
                        _IP.StableTopCount++;
                    }
                    else
                    {
                        _IP.UnStableTopCount++;
                    }
                }
                else if (toeMode == "Sliding")
                {
                    _IP.SlideCount++;
                    if (fofSSliding is double && (double)fofSSliding > 1)
                    {
                        _IP.StableSlideCount++;
                    }
                    else
                    {
                        _IP.UnstableSlideCount++;
                    }
                }
                else
                {
                    _IP.InvalidCount++; // Infinite safety factor
                }

                // Print information for each dynamically feasible iteration
                if (worksheet2 != null)
                {
                    worksheet2.Cells[ini + 2, 12].Value = blockCount;
                    worksheet2.Cells[ini + 2, 13].Value = numStableBlock;
                    worksheet2.Cells[ini + 2, 14].Value = blockCount < 2 ? "-" : failureMode[failureMode.Count - 3];
                    worksheet2.Cells[ini + 2, 15].Value = pn[pn.Count - 2];
                    worksheet2.Cells[ini + 2, 16].Value = rnArray[rnArray.Count - 2];
                    worksheet2.Cells[ini + 2, 17].Value = snArray[snArray.Count - 2];
                    worksheet2.Cells[ini + 2, 18].Value = resistingF;
                    worksheet2.Cells[ini + 2, 19].Value = drivingF;
                    worksheet2.Cells[ini + 2, 20].Value = fofSSliding;
                    worksheet2.Cells[ini + 2, 21].Value = resistingM;
                    worksheet2.Cells[ini + 2, 22].Value = drivingM;
                    worksheet2.Cells[ini + 2, 23].Value = fofSToppling;
                    worksheet2.Cells[ini + 2, 24].Value = toeMode;
                    worksheet2.Cells[ini + 2, 25].Value = toeFS;
                }

                printIteration = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in toe block analysis calculation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void textBox11_TextChanged(object? sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object? sender, EventArgs e)
        {

        }

        /// <summary>
        /// Calculate failure probability
        /// </summary>
        private void CalcProb()
        {
            // *******************************************************************************************
            // Calculate failure probability
            // *******************************************************************************************
            // Get NumTrials, SlideCount, ToppleCount, InvalidCount from Monte Carlo loop and calculate failure probability
            try
            {
                // Define variables
                double probKineBlockTop, probKineFlex, probKineSlide, probKineSlideTop;
                double probSlideFail, probTopFail, totalProbFailure;
                double probKinetic;
                double unstableCount, stableCount;
                double bin = 0, binMax, binSize;
                object foSMean = null!, foSMedian = null!, foSStdDev;
                int icount = 0;

                // Kinematic probability of block toppling
                probKineBlockTop = (double)_IP.KineBlockTopCount / _IP.NumTrials;

                // Kinematic probability of toppling on Set B only (flexural/block flexural)
                probKineFlex = (double)_IP.KineFlexCount / _IP.NumTrials;

                // Kinematic probability of sliding on Set A only
                probKineSlide = (double)_IP.KineSlideCount / _IP.NumTrials;

                // Kinematic probability of sliding and toppling
                probKineSlideTop = (double)_IP.KineSlideTopCount / _IP.NumTrials;

                // Dynamic failure probability (sliding mode), given that block toppling is feasible
                if (_IP.KineBlockTopCount > 0 && _IP.SlideCount > 0)
                {
                    probSlideFail = (double)_IP.UnstableSlideCount / _IP.KineBlockTopCount;
                }
                else
                {
                    probSlideFail = 0;
                }

                // Dynamic failure probability (toppling mode), given that block toppling is feasible
                if (_IP.KineBlockTopCount > 0 && _IP.ToppleCount > 0)
                {
                    probTopFail = (double)_IP.UnStableTopCount / _IP.KineBlockTopCount;
                }
                else
                {
                    probTopFail = 0;
                }

                // Dynamic failure probability (toppling + sliding), given kinematic feasibility
                unstableCount = _IP.UnStableTopCount + _IP.UnstableSlideCount;
                stableCount = _IP.StableTopCount + _IP.StableSlideCount + _IP.InvalidCount;

                if (_IP.KineBlockTopCount > 0)
                {
                    probKinetic = unstableCount / _IP.KineBlockTopCount;
                }
                else
                {
                    probKinetic = 0;
                }

                // Total failure probability
                totalProbFailure = unstableCount / _IP.NumTrials;

                // Create or get frequency worksheet
                ExcelWorksheet freqSheet;
                freqSheet = package.Workbook.Worksheets["frequency"];

                // If block toppling is kinematically feasible
                if (_IP.KineBlockTopCount > 0)
                {
                    // Define frequency distribution and histogram
                    var analysisSheet = package.Workbook.Worksheets["Analysis Details 2"];

                    analysisSheet.Calculate();

                    ExcelRange fosRange = analysisSheet.Cells["AB3:AB20000"];

                    List<double> fosValues = new List<double>();

                    if (fosRange.Any(c => c.Value != null))
                    {
                        foreach (var cell in fosRange)
                        {
                            if (cell.Value != null && cell.Value is double)
                            {
                                fosValues.Add((double)cell.Value);
                            }
                        }

                        if (fosValues.Count > 0)
                        {
                            double mean = fosValues.Average();
                            foSMean = mean;

                            List<double> sortedValues = new List<double>(fosValues);
                            sortedValues.Sort();
                            if (sortedValues.Count % 2 == 0)
                            {
                                foSMedian = (sortedValues[sortedValues.Count / 2 - 1] + sortedValues[sortedValues.Count / 2]) / 2;
                            }
                            else
                            {
                                foSMedian = sortedValues[sortedValues.Count / 2];
                            }

                            // Calculate standard deviation
                            double sumOfSquaresOfDifferences = fosValues.Select(val => (val - mean) * (val - mean)).Sum();
                            foSStdDev = Math.Sqrt(sumOfSquaresOfDifferences / fosValues.Count);

                            if ((double)foSMean > 0)
                            {
                                binMax = Math.Ceiling((double)foSMean + (4 * (double)foSStdDev));
                                binSize = binMax / (binMax * 10);

                                do
                                {
                                    freqSheet.Cells[icount + 2, 1].Value = bin;
                                    bin += binSize;
                                    icount++;
                                } while (bin <= binMax);
                            }
                            else
                            {
                                freqSheet.Cells[icount + 2, 1].Value = "-";
                            }
                        }
                        else
                        {
                            foSMean = "inf.";
                            foSMedian = "inf.";
                            foSStdDev = "inf.";
                        }
                    }
                }

                freqSheet.Cells[2, 1].Value = "";
                freqSheet.Calculate();

                // Print probabilities to "Results" worksheet               
                ExcelWorksheet resultsSheet;
                resultsSheet = package.Workbook.Worksheets["Results"];

                // Main values
                resultsSheet.Cells[5, 6].Value = probKineBlockTop;
                resultsSheet.Cells[7, 6].Value = probKinetic;
                resultsSheet.Cells[9, 6].Value = totalProbFailure;
                resultsSheet.Cells[11, 6].Value = foSMean;
                resultsSheet.Cells[12, 6].Value = foSMedian;

                // Summary of kinematic probabilities
                resultsSheet.Cells[7, 12].Value = _IP.KineBlockTopCount;
                resultsSheet.Cells[7, 13].Value = probKineBlockTop;
                resultsSheet.Cells[8, 12].Value = _IP.KineSlideCount;
                resultsSheet.Cells[8, 13].Value = probKineSlide;
                resultsSheet.Cells[9, 12].Value = _IP.KineFlexCount;
                resultsSheet.Cells[9, 13].Value = probKineFlex;
                resultsSheet.Cells[10, 12].Value = _IP.KineSlideTopCount;
                resultsSheet.Cells[10, 13].Value = probKineSlideTop;
                resultsSheet.Cells[11, 12].Value = _IP.NumTrials - _IP.KineBlockTopCount -
                    _IP.KineSlideCount - _IP.KineSlideTopCount - _IP.KineFlexCount;
                resultsSheet.Cells[11, 13].Value = 1 - probKineBlockTop - probKineSlide - probKineFlex -
                    probKineSlideTop;
                resultsSheet.Cells[12, 12].Value = _IP.NumTrials;
                resultsSheet.Cells[12, 13].Value = probKineBlockTop + probKineSlide + probKineFlex +
                    (1 - probKineBlockTop - probKineSlide - probKineFlex);

                // Summary of dynamic probabilities
                resultsSheet.Cells[18, 10].Value = _IP.SlideCount;
                resultsSheet.Cells[18, 11].Value = _IP.StableSlideCount;
                resultsSheet.Cells[18, 12].Value = _IP.UnstableSlideCount;
                resultsSheet.Cells[18, 13].Value = probSlideFail;
                resultsSheet.Cells[19, 10].Value = _IP.ToppleCount;
                resultsSheet.Cells[19, 11].Value = _IP.StableTopCount;
                resultsSheet.Cells[19, 12].Value = _IP.UnStableTopCount;
                resultsSheet.Cells[19, 13].Value = probTopFail;
                resultsSheet.Cells[20, 10].Value = _IP.KineBlockTopCount;
                resultsSheet.Cells[20, 11].Value = stableCount;
                resultsSheet.Cells[20, 12].Value = unstableCount;
                resultsSheet.Cells[20, 13].Value = probKinetic;

                // Support summary
                bool addToeSupport = false;
                double supportForce = 0;
                double inclination = 0;
                string supportType;

                if (_IP.AddToeSupport == true)
                {
                    supportType = _IP.NatureOfForceApplication;
                    addToeSupport = true;
                    supportForce = _IP.Magnitude;
                    inclination = _IP.Orientation / Math.PI * 180;
                }

                if (addToeSupport)
                {
                    resultsSheet.Cells[25, 10].Value = "Yes";
                    resultsSheet.Cells[25, 11].Value = supportForce;
                    resultsSheet.Cells[25, 12].Value = inclination;
                    resultsSheet.Cells[25, 13].Value = _IP.NatureOfForceApplication;
                }
                else
                {
                    resultsSheet.Cells[25, 10].Value = "No";
                    resultsSheet.Cells[25, 11].Value = "-";
                    resultsSheet.Cells[25, 12].Value = "-";
                    resultsSheet.Cells[25, 13].Value = "-";
                }

                bool boltBlocksTogether = false;
                if (_IP.BoltBlocksTogether == true)
                {
                    boltBlocksTogether = true;
                }

                if (boltBlocksTogether)
                {
                    resultsSheet.Cells[26, 10].Value = "Yes";
                }
                else
                {
                    resultsSheet.Cells[26, 10].Value = "No";
                }

                resultsSheet.Cells[26, 11].Value = "-";
                resultsSheet.Cells[26, 12].Value = "-";
                resultsSheet.Cells[26, 13].Value = "-";



            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error calculating failure probability: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Write all Monte Carlo simulation results to Excel file at once
        /// </summary>
        /// <param name="results">List of all iteration results</param>
        private void WriteAllMonteCarloResultsToExcel(List<MonteCarloIterationResult> results, List<int> safeNums)
        {
            try
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Analysis Details 2"];
                const double Pi = Math.PI;
                // Write all data at once
                for (int i = 0; i < results.Count; i++)
                {
                    MonteCarloIterationResult result = results[i];
                    int row = i + 3; // Start from row 3
                    worksheet.Cells[row, 1].Value = result.IterationNumber;
                    worksheet.Cells[row, 2].Value = Math.Round(result.DipA * (180 / Pi), 0);
                    worksheet.Cells[row, 3].Value = Math.Round(result.DipDirA * (180 / Pi), 0);
                    worksheet.Cells[row, 4].Value = Math.Round(result.DipB * (180 / Pi), 0);
                    worksheet.Cells[row, 5].Value = Math.Round(result.DipDirB * (180 / Pi), 0);
                    worksheet.Cells[row, 6].Value = Math.Round(result.FricA * (180 / Pi), 0);
                    worksheet.Cells[row, 7].Value = Math.Round(result.FricB * (180 / Pi), 0);
                    worksheet.Cells[row, 8].Value = result.KineBlockTop;
                    worksheet.Cells[row, 9].Value = result.KineSlide;
                    worksheet.Cells[row, 10].Value = result.KineFlex;
                    worksheet.Cells[row, 11].Value = result.KineSlideTop;
                }
                for (int i = 0; i < safeNums.Count; i++)
                {
                    for (int k = 12; k <= 25; k++)
                    {
                        worksheet.Cells[safeNums[i] + 2, k].Value = "-";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to Details 2 file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void OpenExcelFile()
        {
            ifPerformAnalysis();

        }
        /// <summary>
        /// Histogram processing method - Adjust output histogram scale
        /// This is a C# implementation of the VB macro "Histo"
        /// </summary>
        private void UpdateHistogram()
        {
            try
            {
                // Get Frequency worksheet
                ExcelWorksheet frequencySheet;
                frequencySheet = package.Workbook.Worksheets["Frequency"];

                // Get Results worksheet
                ExcelWorksheet resultsSheet;
                resultsSheet = package.Workbook.Worksheets["Results"];

                // Get data range
                int lastRowA = 3; // Initial value

                // Find last row in column A of Frequency worksheet
                for (int row = 3; row <= 65536; row++)
                {
                    if (frequencySheet.Cells[row, 1].Value == null)
                    {
                        lastRowA = row - 1;
                        break;
                    }
                }

                // Get X and Y data ranges
                ExcelRange xRange = frequencySheet.Cells[3, 1, lastRowA, 1]; // Column A
                ExcelRange y1Range = frequencySheet.Cells[3, 2, lastRowA, 2]; // Column B
                ExcelRange y2Range = frequencySheet.Cells[3, 4, lastRowA, 4]; // Column D
                ExcelRange y3Range = frequencySheet.Cells[3, 11, lastRowA, 11]; // Column K

                // Get Chart 8 from Results worksheet               
                ExcelChart chart = (ExcelChart)resultsSheet.Drawings[0];

                ExcelChart CumulativeProbability = chart.PlotArea.ChartTypes[1];
                CumulativeProbability.Series[0].Series = y3Range.FullAddress;

                // Update chart data
                if (chart.Series.Count >= 2)
                {
                    // Update chart series data
                    chart.Series[0].XSeries = xRange.FullAddress;
                    chart.Series[0].Series = y1Range.FullAddress;
                    chart.Series[1].Series = y2Range.FullAddress;
                }

                chart.AdjustPositionAndSize();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating histogram: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region File Interaction Module
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // Create parameter object
                var parameters = new FormParameters
                {
                    MainFormParameters = new MainFormParameters
                    {
                        SlopeHeight = double.Parse(textBox1_1.Text),
                        SlopeAngle = double.Parse(textBox1_2.Text),
                        TopAngle = double.Parse(textBox1_3.Text),
                        SlopeDipDir = double.Parse(textBox1_4.Text),
                        MeanDipA = double.Parse(textBox2_11.Text),
                        MeanDipB = double.Parse(textBox2_21.Text),
                        MeanDipDirA = double.Parse(textBox2_12.Text),
                        MeanDipDirB = double.Parse(textBox2_22.Text),
                        MeanSpaceA = double.Parse(textBox3_11.Text),
                        MeanSpaceB = double.Parse(textBox3_21.Text),
                        MeanFricA = double.Parse(textBox3_31.Text),
                        MeanFricB = double.Parse(textBox3_41.Text),
                        MeanSeis = double.Parse(textBox4_11.Text),
                        PorePress = textBox4_21.Text.EndsWith('%') ? double.Parse(textBox4_21.Text.TrimEnd('%')) * 0.01 : double.Parse(textBox4_21.Text),
                        UnitWeight = double.Parse(textBox3_51.Text),
                        UnitWeightH2O = double.Parse(textBox3_61.Text),
                        DistSpaceA = comboBox1.Text,
                        DistSpaceB = comboBox2.Text,
                        DistFricA = comboBox3.Text,
                        DistFricB = comboBox4.Text,
                        DistSeis = comboBox5.Text,
                        DistPorePress = comboBox6.Text,
                        DistUnitWeight = comboBox7.Text,
                        DistUnitWeightH2O = comboBox8.Text,
                        FisherKA = double.Parse(textBox2_13.Text),
                        FisherKB = double.Parse(textBox2_23.Text),
                        StDevSpaceA = double.Parse(textBox3_13.Text),
                        StDevSpaceB = double.Parse(textBox3_23.Text),
                        StDevFricA = double.Parse(textBox3_33.Text),
                        StDevFricB = double.Parse(textBox3_43.Text),
                        StDevUnitWt = double.Parse(textBox3_53.Text),
                        StDevUnitWtH20 = double.Parse(textBox3_63.Text),
                        StDevSeis = double.Parse(textBox4_13.Text),
                        StDevPorePress = textBox4_23.Text,
                        BoltBlocksTogether = checkBox2.Checked,
                        AddToeSupport = checkBox1.Checked
                    },
                    AddSupportParameters = new AddSupportParameters
                    {
                        NatureOfForceApplication = _IP.NatureOfForceApplication,
                        Magnitude = _IP.Magnitude,
                        Orientation = _IP.Orientation * 180 / Math.PI,
                        OptimumOrientationAgainstSliding = _IP.OptimumOrientationAgainstSliding * 180 / Math.PI,
                        OptimumOrientationAgainstToppling = _IP.OptimumOrientationAgainstToppling * 180 / Math.PI,
                        EffectiveWidth = _IP.EffectiveWidth
                    }
                };

                // Get Resources folder path
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string rootPath = Path.GetFullPath(Path.Combine(exePath, @"..\..\..\..\..\"));
                string resourcePath = Path.Combine(rootPath, "Resources");

                // Create Resources directory if it doesn't exist
                if (!Directory.Exists(resourcePath))
                {
                    Directory.CreateDirectory(resourcePath);
                }

                // Create save file dialog
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.DefaultExt = "json";
                    saveFileDialog.InitialDirectory = resourcePath;
                    saveFileDialog.FileName = $"parameters_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.json";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Serialize parameters to JSON and save to file
                        var options = new JsonSerializerOptions { WriteIndented = true };
                        string jsonString = JsonSerializer.Serialize(parameters, options);
                        File.WriteAllText(saveFileDialog.FileName, jsonString);
                        MessageBox.Show("Parameters have been successfully saved to file!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred while saving parameters: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                // Get Resources folder path
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string rootPath = Path.GetFullPath(Path.Combine(exePath, @"..\..\..\..\..\"));
                string resourcePath = Path.Combine(rootPath, "Resources");

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.InitialDirectory = resourcePath;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string jsonString = File.ReadAllText(openFileDialog.FileName);
                        var parameters = JsonSerializer.Deserialize<FormParameters>(jsonString);

                        if (parameters?.MainFormParameters != null)
                        {
                            // Load main form parameters
                            var main = parameters.MainFormParameters;
                            textBox1_1.Text = main.SlopeHeight.ToString();
                            textBox1_2.Text = main.SlopeAngle.ToString();
                            textBox1_3.Text = main.TopAngle.ToString();
                            textBox1_4.Text = main.SlopeDipDir.ToString();
                            textBox2_11.Text = main.MeanDipA.ToString();
                            textBox2_21.Text = main.MeanDipB.ToString();
                            textBox2_12.Text = main.MeanDipDirA.ToString();
                            textBox2_22.Text = main.MeanDipDirB.ToString();
                            textBox3_11.Text = main.MeanSpaceA.ToString();
                            textBox3_21.Text = main.MeanSpaceB.ToString();
                            textBox3_31.Text = main.MeanFricA.ToString();
                            textBox3_41.Text = main.MeanFricB.ToString();
                            textBox4_11.Text = main.MeanSeis.ToString();
                            textBox4_21.Text = (main.PorePress * 100).ToString() + "%";
                            textBox3_51.Text = main.UnitWeight.ToString();
                            textBox3_61.Text = main.UnitWeightH2O.ToString();

                            if (!string.IsNullOrEmpty(main.DistSpaceA)) comboBox1.Text = main.DistSpaceA;
                            if (!string.IsNullOrEmpty(main.DistSpaceB)) comboBox2.Text = main.DistSpaceB;
                            if (!string.IsNullOrEmpty(main.DistFricA)) comboBox3.Text = main.DistFricA;
                            if (!string.IsNullOrEmpty(main.DistFricB)) comboBox4.Text = main.DistFricB;
                            if (!string.IsNullOrEmpty(main.DistSeis)) comboBox5.Text = main.DistSeis;
                            if (!string.IsNullOrEmpty(main.DistPorePress)) comboBox6.Text = main.DistPorePress;
                            if (!string.IsNullOrEmpty(main.DistUnitWeight)) comboBox7.Text = main.DistUnitWeight;
                            if (!string.IsNullOrEmpty(main.DistUnitWeightH2O)) comboBox8.Text = main.DistUnitWeightH2O;

                            textBox2_13.Text = main.FisherKA.ToString();
                            textBox2_23.Text = main.FisherKB.ToString();
                            textBox3_13.Text = main.StDevSpaceA.ToString();
                            textBox3_23.Text = main.StDevSpaceB.ToString();
                            textBox3_33.Text = main.StDevFricA.ToString();
                            textBox3_43.Text = main.StDevFricB.ToString();
                            textBox3_53.Text = main.StDevUnitWt.ToString();
                            textBox3_63.Text = main.StDevUnitWtH20.ToString();
                            textBox4_13.Text = main.StDevSeis.ToString();
                            textBox4_23.Text = main.StDevPorePress;
                            checkBox2.Checked = main.BoltBlocksTogether;
                            checkBox1.Checked = main.AddToeSupport;
                        }

                        if (parameters?.AddSupportParameters != null)
                        {
                            // Load support parameters
                            var support = parameters.AddSupportParameters;
                            _IP.NatureOfForceApplication = support.NatureOfForceApplication ?? "Passive";
                            _IP.Magnitude = support.Magnitude;
                            _IP.Orientation = support.Orientation * (Math.PI / 180);  // Convert degrees to radians
                            _IP.OptimumOrientationAgainstSliding = support.OptimumOrientationAgainstSliding * (Math.PI / 180);  // Convert degrees to radians
                            _IP.OptimumOrientationAgainstToppling = support.OptimumOrientationAgainstToppling * (Math.PI / 180);  // Convert degrees to radians
                            _IP.EffectiveWidth = support.EffectiveWidth;
                        }

                        MessageBox.Show("Parameters have been successfully imported!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred while importing parameters: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred while opening Excel file: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

    }

}
