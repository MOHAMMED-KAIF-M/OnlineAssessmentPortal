# Question Quality Report

- Workbooks scanned: 1081
- Proper workbooks: 929
- Not proper workbooks: 152

## Audit Rules

- `needs_rebuild(...)` flags: wrong question count, wrong raw row count, non-structured sheet, blank/unsupported/nonstandard type, blank question
- Question-level flags: MCQ rows without 4 split options, packed question rows, dataset rows inside question sheets, leaked header rows, duplicate question text

## Top Issues

- `mcq_missing_split_options`: 797 occurrences across 144 workbooks
- `packed_question_row`: 204 occurrences across 36 workbooks
- `dataset_row_in_questions_sheet`: 204 occurrences across 36 workbooks
- `duplicate_question_text`: 42 occurrences across 32 workbooks
- `count_19`: 36 occurrences across 36 workbooks
- `raw_count_19`: 36 occurrences across 36 workbooks
- `count_26`: 4 occurrences across 4 workbooks
- `raw_count_26`: 4 occurrences across 4 workbooks
- `scenario_partial_options`: 2 occurrences across 2 workbooks
- `count_27`: 1 occurrences across 1 workbooks
- `raw_count_27`: 1 occurrences across 1 workbooks
- `count_29`: 1 occurrences across 1 workbooks
- `raw_count_29`: 1 occurrences across 1 workbooks

## Worst Affected Workbooks

| Status | Source File | Workbook Title | Question Count | Issues |
|---|---|---|---:|---|
| not_proper | 0336.xlsx | . Product Pricing with Prescriptive Modeling | 19 | packed_question_row=7, dataset_row_in_questions_sheet=7, mcq_missing_split_options=7, count_19=1, raw_count_19=1 |
| not_proper | 0337.xlsx | • AssignmentKERC Inc, Optimum Manufacturing Quantity | 19 | packed_question_row=7, dataset_row_in_questions_sheet=7, mcq_missing_split_options=7, count_19=1, raw_count_19=1 |
| not_proper | 0341.xlsx | Advanced Functions VLOOKUP, XLOOKUP, IF, INDEX-MATCH, etc | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0506.xlsx | Back propagation, Gradient Descent | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0512.xlsx | . Hands-on Decision Tree with ML Tool | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0516.xlsx | . Hands-on KNN with ML Tool | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0519.xlsx | . Hands-on Linear Regression with ML Tool | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0522.xlsx | . Hands-on Logistics Regression with ML Tool | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0525.xlsx | . Hands-onSVM with ML Tool | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, duplicate_question_text=2, count_19=1 |
| not_proper | 0339.xlsx | . Hands on Regression Modeling in Excel | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0340.xlsx | . Mathematics behind Linear Regression | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0342.xlsx | Advanced Pivot Table Techniques | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0507.xlsx | . Clustering, Classification And Regression | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0508.xlsx | . ML Workflow, Popular ML Algorithms | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0509.xlsx | . Supervised Vs Unsupervised | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0510.xlsx | . What Is MLML Vs AI | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0513.xlsx | . Hands-onK Means Clustering | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0515.xlsx | Introduction to KMeans and How it works | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0517.xlsx | . Introduction to KNN_ Nearest Neighbor | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0518.xlsx | . Regression with KNN | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0520.xlsx | . How it works Regression and Best Fit Line | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0521.xlsx | . Introduction to Linear Regression | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0523.xlsx | . Introduction to Logistic Regression_ | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0524.xlsx | Classification & Sigmoid Curve | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |
| not_proper | 0335.xlsx | . Predictive Analytics with Low Uncertainty_ Case Study | 19 | packed_question_row=6, dataset_row_in_questions_sheet=6, mcq_missing_split_options=6, count_19=1, raw_count_19=1 |

## Sample Problem Rows

### 0336.xlsx - . Product Pricing with Prescriptive Modeling
- packed_question_row: SECTION C: Dataset,Dataset,Price ($),,Quantity Demanded,Revenue ($),Cost per Unit ($),Total Cost ($),Profit ($)
- dataset_row_in_questions_sheet: SECTION C: Dataset,Dataset,Price ($),,Quantity Demanded,Revenue ($),Cost per Unit ($),Total Cost ($),Profit ($)
- mcq_missing_split_options: SECTION C: Dataset,Dataset,Price ($),,Quantity Demanded,Revenue ($),Cost per Unit ($),Total Cost ($),Profit ($)
- packed_question_row: SECTION C: Dataset,Dataset,20,,500,,12,,
- dataset_row_in_questions_sheet: SECTION C: Dataset,Dataset,20,,500,,12,,

### 0337.xlsx - • AssignmentKERC Inc, Optimum Manufacturing Quantity
- packed_question_row: SECTION C: Dataset,Dataset,Parameter,,Value,,,,,
- dataset_row_in_questions_sheet: SECTION C: Dataset,Dataset,Parameter,,Value,,,,,
- mcq_missing_split_options: SECTION C: Dataset,Dataset,Parameter,,Value,,,,,
- packed_question_row: SECTION C: Dataset,Dataset,Annual Demand (D),,12000,,,,,
- dataset_row_in_questions_sheet: SECTION C: Dataset,Dataset,Annual Demand (D),,12000,,,,,

### 0341.xlsx - Advanced Functions VLOOKUP, XLOOKUP, IF, INDEX-MATCH, etc
- duplicate_question_text: Which function?
- duplicate_question_text: Which function?
- duplicate_question_text: Function?
- duplicate_question_text: Function?
- duplicate_question_text: Function?

### 0506.xlsx - Back propagation, Gradient Descent
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Effect?
- duplicate_question_text: Effect?
- packed_question_row: SECTION C: Dataset,Dataset,Predicted,,Actual,Error,,,

### 0512.xlsx - . Hands-on Decision Tree with ML Tool
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?

### 0516.xlsx - . Hands-on KNN with ML Tool
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?

### 0519.xlsx - . Hands-on Linear Regression with ML Tool
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?

### 0522.xlsx - . Hands-on Logistics Regression with ML Tool
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?

### 0525.xlsx - . Hands-onSVM with ML Tool
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?

### 0339.xlsx - . Hands on Regression Modeling in Excel
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?
- duplicate_question_text: Purpose?
- packed_question_row: SECTION C: Dataset,Dataset,X,,Y,,,,
- dataset_row_in_questions_sheet: SECTION C: Dataset,Dataset,X,,Y,,,,

### 0340.xlsx - . Mathematics behind Linear Regression
- duplicate_question_text: what does it mean?
- duplicate_question_text: What does it mean?
- packed_question_row: SECTION C: Dataset,Dataset,X,,Y,,,,
- dataset_row_in_questions_sheet: SECTION C: Dataset,Dataset,X,,Y,,,,
- mcq_missing_split_options: SECTION C: Dataset,Dataset,X,,Y,,,,

### 0342.xlsx - Advanced Pivot Table Techniques
- duplicate_question_text: Tool?
- duplicate_question_text: Tool?
- duplicate_question_text: Tool?
- duplicate_question_text: Tool?
- packed_question_row: SECTION C: Dataset,Dataset,Region,,Product,Sales,,

### 0507.xlsx - . Clustering, Classification And Regression
- duplicate_question_text: Type?
- duplicate_question_text: Type?
- duplicate_question_text: Type?
- duplicate_question_text: Type?
- duplicate_question_text: Type?

### 0508.xlsx - . ML Workflow, Popular ML Algorithms
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- duplicate_question_text: Step?
- packed_question_row: SECTION C: Dataset,Dataset,Feature,,Label,,,,

### 0509.xlsx - . Supervised Vs Unsupervised
- duplicate_question_text: Type?
- duplicate_question_text: Type?
- duplicate_question_text: Type?
- duplicate_question_text: Type?
- duplicate_question_text: Type?
