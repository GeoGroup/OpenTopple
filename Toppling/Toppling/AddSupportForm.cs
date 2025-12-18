namespace Toppling
{
    public partial class AddSupportForm : Form
    {
        private readonly double _meanDipA;
        private readonly double _meanFricA;

        // Add properties for parameter passing, set default values
        public string NatureOfForceApplication { get; private set; } = string.Empty;
        public double Magnitude { get; private set; }
        public double Orientation { get; private set; }
        public double OptimumOrientationAgainstSliding { get; private set; }
        public double OptimumOrientationAgainstToppling { get; private set; }
        public double EffectiveWidth { get; private set; }

        public AddSupportForm(double meanDipA, double meanFricA)
        {
            InitializeComponent();
            _meanDipA = meanDipA;
            _meanFricA = meanFricA;
            InitializeDefaultValues();
            
            // Set textbox3 and textbox4 as read-only
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            
            // Calculate and set initial values
            UpdateOptimumOrientations();
        }

        // Add constructor with parameters
        public AddSupportForm(string natureOfForceApplication, double magnitude, double orientation, 
            double optimumOrientationAgainstSliding, double optimumOrientationAgainstToppling, double effectiveWidth,
            double meanDipA, double meanFricA)
        {
            InitializeComponent();
            _meanDipA = meanDipA;
            _meanFricA = meanFricA;
            
            // Set control values
            comboBox9.Text = natureOfForceApplication;
            textBox1.Text = magnitude.ToString();
            textBox2.Text = orientation.ToString();
            comboBox10.Text = effectiveWidth.ToString();

            // Set textbox3 and textbox4 as read-only
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            
            // Calculate and set initial values
            UpdateOptimumOrientations();
        }

        private void InitializeDefaultValues()
        {
            comboBox9.Text = "Passive";
            textBox1.Text = "175";
            textBox2.Text = "6";
            comboBox10.Text = "2";
        }

        private void UpdateOptimumOrientations()
        {
            // Calculate OptimumOrientationAgainstSliding: MeanDipA - MeanFricA
            double optimumOrientationSliding = _meanDipA - _meanFricA;
            textBox3.Text = optimumOrientationSliding.ToString();

            // Calculate OptimumOrientationAgainstToppling: -MeanDipA
            double optimumOrientationToppling = -_meanDipA;
            textBox4.Text = optimumOrientationToppling.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate input
                if (string.IsNullOrEmpty(comboBox9.Text) || 
                    string.IsNullOrEmpty(textBox1.Text) ||
                    string.IsNullOrEmpty(textBox2.Text) ||
                    string.IsNullOrEmpty(textBox3.Text) ||
                    string.IsNullOrEmpty(textBox4.Text) ||
                    string.IsNullOrEmpty(comboBox10.Text))
                {
                    MessageBox.Show("Please fill in all required parameters!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Save parameters
                NatureOfForceApplication = comboBox9.Text;
                Magnitude = double.Parse(textBox1.Text);
                Orientation = double.Parse(textBox2.Text) * (Math.PI / 180);
                OptimumOrientationAgainstSliding = double.Parse(textBox3.Text) * (Math.PI / 180);
                OptimumOrientationAgainstToppling = double.Parse(textBox4.Text) * (Math.PI / 180);
                EffectiveWidth = double.Parse(comboBox10.Text);

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Input parameter format error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
