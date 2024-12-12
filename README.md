# Batch Processing of Excel Files in Multiple Folders

## Overview

This script automates the processing of Excel files located in multiple subfolders (CT1, CT2, ..., CT15). Each subfolder contains Excel files with a sheet named Detail_5_2_X_0_1 where X is a number from 1 to 10 (corresponding to experimental tracks).

The script reads the specified sheets, processes them, and outputs two Excel files per subfolder:

    CT[i]_stitched.xlsx
    CT[i]_extract.xlsx

## Problem Statement

Currently, the script processes Excel files in a single folder (e.g., CT1) and generates the required outputs. However, extending this logic to loop through all CT[i] folders requires additional modifications. This README outlines the necessary steps to build the loop and explains the solution.

## Requirements

    Python installed with the following libraries:
        pandas
        glob
        os

## Outputs

    CT[i]_stitched.xlsx: Contains combined data from all relevant sheets.
    CT[i]_extract.xlsx: Contains a smaller, extracted portion of the data.

# Notes

    Ensure all Excel files follow the expected naming and sheet structure.
    Modify the data processing logic as needed to match your requirements.
    Handle errors gracefully to debug issues with specific files or sheets.
