# Survey Analytics Automation App

# Overview
A Streamlit-based application that automates the end-to-end survey analytics workflow, including scripting, recoding, weighting, toplines, and crosstabs. Designed to replace manual Excel and SPSS processes, the app improves speed, accuracy, and consistency in survey reporting.

# Impact
- Reduced manual processing time by ~50%
- Eliminated common copy-and-paste errors
- Enabled non-technical users to generate complex analytical outputs

# Purpose

The purpose of this project is to streamline the polling workflow and reduce human error by replacing manual Excel and SPSS processes with a single, structured application.
By automating scripting, recodes, weighting, toplines, and crosstabs, the app minimizes copy-and-paste mistakes, ensures consistency across outputs, and makes complex survey production faster and more reliable.

# Modules
Module 0 – Project

Handles saving and loading complete projects using JSON files.
This allows work to be paused, resumed, and shared without losing state, while maintaining backward compatibility across app updates.

# Module 1 – Scripting

Imports survey scripts from Excel or text and parses them into structured variables.
This module detects question headers, batteries, and value labels so that questions are defined once and reused consistently throughout the workflow.

# Module 2 – Recodes

Creates SPSS-style recoded variables by grouping response options from existing questions.
Supports ELSE=COPY behavior to preserve unused categories and ensures recodes stay synchronized with the original variables.

# Module 2.5 – Derived Variables

Defines new variables using metadata-based rules without requiring a dataset.
This allows transformations and logic to be built early and applied consistently once data is loaded.

# Module 3 – Import & Match Data

Loads raw survey data and matches it against the scripted variable catalog.
This ensures variable names, labels, and codes align correctly between the script and the dataset.

# Module 4 – Weighting

Builds and applies survey weighting schemes.
Tracks weighting stages, stores weighting configurations, and generates reproducible SPSS syntax.

# Module 5 – Topline Shell

Creates topline templates that define question order and optional injected recodes.
Serves as the structural blueprint for topline reporting.

# Module 6 – Weighted Toplines

Generates weighted toplines using correct denominators, including questions answered by subsets of respondents.
Ensures percentages and totals are calculated accurately and consistently.

# Module 7 – Crosstabs

Produces weighted crosstabs across selected variables.
Optimized for large surveys with canonical variable matching, fuzzy resolution, and performance-conscious UI design.
