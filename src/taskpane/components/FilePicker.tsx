import * as React from "react";
import {  FilePicker, IFilePickerResult } from "ulteam-react";
import { IFilePickerProps } from "ulteam-react/lib/common/components/FilePicker/IFilePickerProps";

const onChangeFilePicker = (files: File[]) => {
  for (const index in files) {
    console.log("file selected", files[index]);
  }
};

const onChangeFilePickerAsBase64 = (files: IFilePickerResult[]) => {
  for (const index in files) {
    if (index) {
      console.log(`base64: ${files[index].base64}`);
    }
  }
};

export const csvPickerProps: IFilePickerProps = {
  accumulateFiles: true,
  errorMessage: "Please pick a CSV file",
  label: "File picker label",
  text: "Browse...",
  multiple: true,
  onChangeFile: onChangeFilePicker,
  onChangeFileBase64: onChangeFilePickerAsBase64,
  placeholder: "Please pick a CSV file",
  required: true
};

export const CSVPicker = (p: IFilePickerProps = csvPickerProps) => {
  //let { label, multiple } = p;
  console.log(p)
  return (
    <FilePicker></FilePicker>
  );
};
