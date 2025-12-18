# OpenTopple
OpenTopple is a C# WinForms open-source application for probabilistic block-toppling analysis in rock slopes. Rewritten from ROCKTOPPLE (Excel/VBA), it features a modern UI, interactive 2D visualization, and high-performance Monte Carlo simulation without requiring Excel. Supports parameter import/export and deterministic/probabilistic analysis.

## Project Structure
```
Toppling/
├── Resources/    # Resource files directory
├── Results/    # Results output directory
└── Toppling.sln    # Solution file
```

## System Requirements
- Windows 10/11
- .NET 8.0
  
## Usage Instructions
1. Basic Parameter Setup   
   - Set basic slope parameters (height, angle, etc.)
   - Configure joint set parameters (dip direction, dip angle, etc.)
   - Set physical parameters (friction angle, unit weight, etc.)

2. Analysis Setup
   - Select parameter distribution types (Normal, Log-normal, Exponential)
   - Set Monte Carlo simulation iterations
   - Configure support system parameters (if needed)

3. Run Analysis
   - Click "Preview Geometry" to check parameter settings
   - Execute Monte Carlo simulation
   - View analysis results and charts

4. Result Processing
   - Export analysis charts
   - Analyze failure probability
   - Review generated Excel reports

## License
MIT

## Contact Information
zya@hhu.edu.cn 
