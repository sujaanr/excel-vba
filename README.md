# Excel-VBA

# Excel VBA Chart Macro

This VBA macro automatically generates a clustered bar chart based on data within a specified range in an active Excel worksheet. It customizes the chart by setting its size, data range, type, adding a title, legend, and formatting various elements for clarity and visual appeal.

## Features

- **Automatic Chart Generation:** Creates a clustered bar chart with minimal input.
- **Custom Formatting:** Applies specific formatting to the chart title, axes, and legend.
- **Flexible Data Range:** While preset to a specific range, it can be modified to suit different data sets.
- **No Gridlines:** Removes clutter by hiding gridlines, focusing attention on the data.

## Requirements

- **Microsoft Excel:** Ensure you have a version of Excel that supports VBA (most versions do).
- **Macro Settings:** Macros must be enabled in Excel to run this script.

## Setup Instructions

1. **Open Excel:** Navigate to the workbook where you want to add the chart.
2. **Enable Developer Tab:** If not already enabled, go to `File` > `Options` > `Customize Ribbon` and check the `Developer` checkbox.
3. **Open VBA Editor:** On the Developer tab, click on `Visual Basic`, or press `Alt` + `F11`.
4. **Insert a Module:** In the VBA editor, right-click on any of the objects for your workbook listed in the Project Explorer. Select `Insert` > `Module`.
5. **Copy the Macro:** Paste the provided VBA code into the newly created module window.
6. **Modify Data Range (Optional):** Adjust the `DataRng` variable if your data is in a different range than the default `B3:E7`.
7. **Run the Macro:** Press `F5` to run the macro, or close the VBA editor and run the macro from the `Macros` option under the Developer tab.

## Usage Notes

- The macro adds the chart to the active sheet; ensure the correct sheet is active before running.
- Adjust the `Chrt.Chart.SetSourceData Source:=DataRng` line to match your data's location.
- The default chart type is a clustered bar chart (`xlBarClustered`). This can be changed to any other Excel chart type by modifying the `Chrt.Chart.ChartType` line.

## Customization

- Chart title, legend, and axes titles can be customized by modifying the respective sections of the VBA code.
- For advanced customization (like chart style, color, etc.), additional VBA properties and methods can be explored and applied.

## Support

For questions or issues with using this macro, consult the Excel VBA documentation or forums dedicated to Excel VBA development.
