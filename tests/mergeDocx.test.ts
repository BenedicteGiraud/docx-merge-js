import { describe, it, expect } from "vitest";
import { mergeDocx } from "../src/index";
import { existsSync, unlinkSync } from "fs";

describe("mergeDocx", () => {
  it("merges two DOCX files and creates output file", () => {
    const file1 = "./tests/fixtures/template.docx";
    const file2 = "./tests/fixtures/table.docx";
    const output = "./tests/fixtures/output.docx";

    mergeDocx(file1, file2, { outputPath: output, pattern: "{{table}}" });

    // Check if the output file exists
    expect(existsSync(output)).toBe(true);

    // cleanup
    unlinkSync(output);
  });

  it("merges two DOCX files and returns buffer", () => {
    const file1 = "./tests/fixtures/template.docx";
    const file2 = "./tests/fixtures/table.docx";

    const buffer = mergeDocx(file1, file2, { pattern: "{{table}}" });

    // Check if the buffer is not empty
    expect(buffer).toBeInstanceOf(Buffer);
  });

  it("throws error if pattern is not found", () => {
    const file1 = "./tests/fixtures/template.docx";
    const file2 = "./tests/fixtures/table.docx";
    const invalidPattern = "{{invalid}}";
    expect(() => mergeDocx(file1, file2, { pattern: invalidPattern })).toThrow(
      "No matching pattern found in the template XML"
    );
  });

  it("throws error if source or content path is missing", () => {
    expect(() => mergeDocx("", "./tests/fixtures/table.docx", {})).toThrow(
      "Missing source or content path"
    );
    expect(() => mergeDocx("./tests/fixtures/template.docx", "", {})).toThrow(
      "Missing source or content path"
    );
  });

  it("throws error if insert position or pattern is missing", () => {
    expect(() =>
      mergeDocx(
        "./tests/fixtures/template.docx",
        "./tests/fixtures/table.docx",
        {}
      )
    ).toThrow("Missing insert position or pattern");
  });

  it("throws error if the path are not valid", () => {
    expect(() =>
      mergeDocx(
        "./tests/fixtures/nonexistent.docx",
        "./tests/fixtures/table.docx",
        { pattern: "{{table}}" }
      )
    ).toThrow("Source file does not exist: ./tests/fixtures/nonexistent.docx");

    expect(() =>
      mergeDocx(
        "./tests/fixtures/template.docx",
        "./tests/fixtures/nonexistent.docx",
        { pattern: "{{table}}" }
      )
    ).toThrow("Content file does not exist: ./tests/fixtures/nonexistent.docx");
  });

  it("throws error if the document are not docx files", () => {
    expect(() =>
      mergeDocx(
        "./tests/fixtures/template.pdf",
        "./tests/fixtures/table.docx",
        { pattern: "{{table}}" }
      )
    ).toThrow("Invalid file extension. Only .docx files are supported");

    expect(() =>
      mergeDocx(
        "./tests/fixtures/template.docx",
        "./tests/fixtures/template.pdf",
        { pattern: "{{table}}" }
      )
    ).toThrow("Invalid file extension. Only .docx files are supported");
  });
});
