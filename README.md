### 概述
这是我为母亲编写的一个程序。她在公司担任财务工作，经常需要处理类似的表格，并逐行对比其中的内容。这个自动化程序能够帮助她节省大量的时间和精力.

### 主要功能
- **缺失数据检测**: 自动检查 `A.xlsx` 和 `B.xlsx` 中的关键列（如 `数量`, `采购单价 含税`, `代码`, `备注`, `价税合计`），输出缺失数据的位置，并提示用户是否继续或退出。
- **Excel数据处理**: 计算 `A.xlsx` 中的总价，并按 `代码` 汇总。随后与 `B.xlsx` 中的数据进行处理并生成汇总结果。
- **生成新文件**: 程序会生成以下两个文件：
    - `TEMP.xlsx`: 包含按 `代码` 汇总的 `A.xlsx` 的总价。
    - `result_new.xlsx`: 包含处理过的 `B.xlsx` 数据，并计算出 `差值` 列。
- **条件高亮**: 在 `A.xlsx` 文件中，程序会高亮所有与 `result_new.xlsx` 中 `差值` 为 0 的 `代码` 行。
- **用户选择**: 如果在 `A.xlsx` 或 `B.xlsx` 中发现缺失数据，用户可以选择输入 `continue` 忽略问题并继续运行，或者按 `Enter` 退出程序以修复数据。
- **程序交互**: 在所有处理完成后，程序会等待用户按下 `Enter` 键以关闭。

### 文件生成
1. **`TEMP.xlsx`**: 由 `A.xlsx` 生成，包含按 `代码` 汇总的 `总价`。
2. **`result_new.xlsx`**: 由 `B.xlsx` 生成，包含原始的 `备注`、`价税合计`，以及程序处理后的 `代码`、`总价` 和 `差值`。
3. **`A_highlighted.xlsx`**: 高亮显示 `差值` 为 0 的代码行。

### 运行环境
- 该程序使用以下 Python 库：
    - `pandas`
    - `openpyxl`
    - `os`
    - `sys`
- 程序兼容 Windows 系统，建议在 Windows 10 或 11 上运行。

### 使用说明
1. 运行程序时，请确保 `A.xlsx` 和 `B.xlsx` 文件与程序位于同一目录中。
2. 程序会自动检测文件中的缺失数据，并提示用户选择是否继续运行或退出。
3. 程序会在当前目录生成 `TEMP.xlsx` 和 `result_new.xlsx` 文件，并更新 `A_highlighted.xlsx`，在其中高亮显示 `差值` 为 0 的行。
4. 程序完成后，会等待用户按下 `Enter` 键关闭。

### 已知问题
- 程序不处理非数值列中的数据异常，例如文本或符号。请确保数据类型正确。

# ENG ver.
### Overview
This is a program I developed for my mother, who works in finance and often needs to manually process and compare rows of data in similar spreadsheets. This automated solution is designed to help her save significant time and effort by streamlining the process.

### Key Features
- **Missing Data Detection**: Automatically checks for missing data in key columns of `A.xlsx` and `B.xlsx` (e.g., `数量`, `采购单价 含税`, `代码`, `备注`, `价税合计`), outputs the missing data locations, and prompts the user to either continue or exit the program.
- **Excel Data Processing**: Calculates the total price in `A.xlsx` and summarizes it by `代码`. It then processes `B.xlsx` and generates a summarized result.
- **File Generation**: The program generates the following two files:
    - `TEMP.xlsx`: Contains the summarized total price from `A.xlsx` by `代码`.
    - `result_new.xlsx`: Contains processed data from `B.xlsx`, along with a calculated `差值` (difference) column.
- **Conditional Highlighting**: In `A.xlsx`, the program highlights all rows where the `代码` matches with rows in `result_new.xlsx` that have a `差值` of 0.
- **User Choice**: If missing data is detected in `A.xlsx` or `B.xlsx`, the user can choose to either type `continue` to ignore the issue and proceed, or press `Enter` to exit and fix the data.
- **Program Interaction**: After completing all operations, the program will wait for the user to press `Enter` to close.

### Generated Files
1. **`TEMP.xlsx`**: Generated from `A.xlsx`, containing the total price summarized by `代码`.
2. **`result_new.xlsx`**: Generated from `B.xlsx`, containing the original `备注` and `价税合计`, along with the processed `代码`, `总价`, and `差值`.
3. **`A_highlighted.xlsx`**: Highlights rows in `A.xlsx` where the `代码` matches rows in `result_new.xlsx` with a `差值` of 0.

### Environment Requirements
- The program requires the following Python libraries:
    - `pandas`
    - `openpyxl`
    - `os`
    - `sys`
- Compatible with Windows systems, recommended for Windows 10 or 11.

### Instructions
1. Ensure that the `A.xlsx` and `B.xlsx` files are in the same directory as the program before running.
2. The program will automatically check for missing data in the files and prompt the user to choose whether to continue or exit.
3. The program will generate `TEMP.xlsx` and `result_new.xlsx` in the current directory, and update `A_highlighted.xlsx` with highlighted rows where the `差值` is 0.
4. After the program completes, it will wait for the user to press `Enter` to close the console.

### Known Issues
- The program does not handle non-numeric data in key columns. Ensure that the data types are correct before running.
