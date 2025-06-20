# Recruitment KPI Report Generator

## Overview
This VBA script automatically generates a professional PowerPoint report analyzing recruitment metrics, including:
- Key hiring statistics
- Position-based analysis
- Recruitment source effectiveness
- Actionable recommendations

## Features
- **Automated PowerPoint generation** - Creates a complete presentation with 5 slides
- **Data visualization** - Includes clustered column and bar charts
- **Professional formatting** - Applies corporate templates (when available)
- **Error handling** - Robust checks for PowerPoint availability and template locations
- **Desktop saving** - Automatically saves to user's desktop with date-stamped filename

## Requirements
- Microsoft PowerPoint (2013 or newer recommended)
- Microsoft Excel (for VBA editor)
- PowerPoint template "Slice.thmx" (optional, falls back to default theme)

## Customization
To adapt the report to your organization:

### Data Customization
1. Modify the `chartData` arrays in the code to reflect your metrics:
   - Position analysis (Slide 3)
   - Source analysis (Slide 4)

2. Update the key metrics table (Slide 2) with your organization's data

### Visual Customization
1. To change the template:
   - Replace "Slice.thmx" with your corporate template name
   - Add additional search paths in the `FindTemplate` function if needed

2. To modify chart styles:
   - Adjust the RGB values in `.ChartArea.Format.Fill.ForeColor.RGB`
   - Change chart types by modifying the `xlColumnClustered`/`xlBarClustered` constants

## Troubleshooting
### Common Issues
1. **PowerPoint won't start**:
   - Ensure PowerPoint is installed
   - Check macro security settings allow VBA to run

2. **Template not found**:
   - Verify template exists in one of the searched paths
   - Or remove template application to use default theme

3. **Charts not displaying**:
   - Check if chart data ranges are correctly defined
   - Verify PowerPoint version supports the chart types used

### Debugging Tips
- Step through the code (`F8` in VBA editor) to identify where errors occur
- Check the immediate window (`CTRL+G`) for debug messages
- Add `MsgBox` statements to verify variable values at key points

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing
Contributions are welcome! Please:
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Open a pull request

## Sample Data Structure
The report uses these data structures:

### Position Analysis
| Position          | Time (Days) | Cost ($) | Performance |
|-------------------|------------|----------|-------------|
| Product Manager   | 11.2       | 328.1    | 0.38        |
| Data Analyst      | 7.3        | 207.0    | 0.19        |

### Source Analysis
| Source         | Time (Days) | Cost ($) | Performance |
|---------------|------------|----------|-------------|
| Referral      | 10.8       | 253.5    | 0.32        |
| Company Site  | 8.8        | 186.6    | 0.21        |

## Version History
- 1.0 (Current) - Initial release with core functionality
- 1.1 (Planned) - Add external data source integration

## Acknowledgments
- Microsoft Documentation for PowerPoint VBA reference
- Stack Overflow community for troubleshooting help

For support or feature requests, please open an issue in the repository.
