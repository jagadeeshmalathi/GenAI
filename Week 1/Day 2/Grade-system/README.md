# Student Grade System

A simple Python project that calculates a student's grade based on marks entered by the user.

This project was created as part of **Week 1 - Day 2** Python practice.

## Project Overview

The program:

- Accepts student marks as input
- Checks the entered value
- Calculates the grade using the conditions written in the notebook
- Displays the final grade to the user

## Project Structure

```text
Grade-system/
├── day2.ipynb
└── README.md
```

## Requirements

- Python 3
- Visual Studio Code
- Jupyter extension for VS Code

## How to Run the Project

### Run in Visual Studio Code

1. Open Visual Studio Code.
2. Open the following folder:

```text
C:\GenAI - Social Eagle\Week 1\Day 2\Grade-system
```

3. Open `day2.ipynb`.
4. Select a Python kernel from the top-right corner.
5. Run each notebook cell using the **Run** button.
6. Enter the requested marks when the input box appears.
7. The calculated grade will be displayed below the cell.

## Example

```text
Enter student marks: 85
Grade: A
```

The exact grade depends on the grading conditions defined inside `day2.ipynb`.

## Troubleshooting

### Python kernel is not available

Install Python and the Jupyter extension, then restart Visual Studio Code.

You can verify Python using:

```powershell
python --version
```

Install Jupyter using:

```powershell
python -m pip install jupyter ipykernel
```

### Input box is not visible

When the notebook uses `input()`, the input field normally appears near the top of the VS Code window or directly below the running cell.

Click inside the input field, enter the value, and press **Enter**.

## Learning Objectives

This project helps practice:

- Python variables
- User input
- Conditional statements
- Comparison operators
- Grade calculation logic
- Running Python code in Jupyter Notebook

## Author

Jagadeesh
