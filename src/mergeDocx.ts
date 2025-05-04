import AdmZip from "adm-zip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import fs from "fs";

const documentXMLPath = "word/document.xml";

const parseXML = (xmlString: string) => {
  const options = {
    allowBooleanAttributes: true,
    attributeNamePrefix: "@_",
    ignoreAttributes: false,
    preserveOrder: true,
  };
  const parser = new XMLParser(options);
  return parser.parse(xmlString);
};

const buildXML = (xmlObject: any) => {
  const options = {
    attributeNamePrefix: "@_",
    format: true,
    ignoreAttributes: false,
    preserveOrder: true,
  };
  const builder = new XMLBuilder(options);
  return builder.build(xmlObject);
};

const mergeParsedXML = (
  templateParsedXML: any,
  contentParsedXMLBody: any,
  {
    pattern,
    insertStart,
    insertEnd,
  }: {
    pattern?: string;
    insertStart?: boolean;
    insertEnd?: boolean;
  }
) => {
  if (!Array.isArray(templateParsedXML)) {
    throw new Error("Template XML is not an array");
  }
  let numberOfReplacements = 0 + (insertStart ? 1 : 0) + (insertEnd ? 1 : 0);
  const mergedParsedXML = templateParsedXML.map((item) => {
    if (!item["w:document"]) {
      return item;
    }

    const documentXml = item["w:document"];

    if (!Array.isArray(documentXml)) {
      throw new Error("<w:document> tag is not an array");
    }
    return {
      ...item,
      "w:document": documentXml.map((docItem) => {
        if (!docItem["w:body"]) {
          return docItem;
        }

        const bodyContent = docItem["w:body"];

        if (!Array.isArray(bodyContent)) {
          throw new Error("<w:body> tag is not an array");
        }
        return {
          ...docItem,
          "w:body": bodyContent.reduce(
            (acc, bodyItem) => {
              if (bodyItem["w:sectPr"]) {
                return [
                  ...acc,
                  ...(insertEnd ? contentParsedXMLBody : []),
                  bodyItem,
                ];
              }

              if (!bodyItem["w:p"] || !pattern) {
                return [...acc, bodyItem];
              }

              if (!Array.isArray(bodyItem["w:p"])) {
                throw new Error("<w:p> tag is not an array");
              }
              const elementIsMatchingPattern = bodyItem["w:p"]?.some((p) => {
                if (!p["w:r"]) {
                  return false;
                }
                if (!Array.isArray(p["w:r"])) {
                  throw new Error("<w:r> tag is not an array");
                }
                return p["w:r"].some((r) => {
                  if (!r["w:t"]) {
                    return false;
                  }
                  if (!Array.isArray(r["w:t"])) {
                    throw new Error("<w:t> tag is not an array");
                  }
                  return r["w:t"].some((t) => t["#text"] === pattern);
                });
              });

              if (!elementIsMatchingPattern) {
                return [...acc, bodyItem];
              }

              numberOfReplacements += 1;
              return [...acc, ...contentParsedXMLBody];
            },
            insertStart ? contentParsedXMLBody : []
          ),
        };
      }),
    };
  });

  if (numberOfReplacements === 0) {
    throw new Error("No matching pattern found in the template XML");
  }

  return mergedParsedXML;
};

const getXmlBody = (xmlArray: any) => {
  if (!Array.isArray(xmlArray)) {
    throw new Error("Input is not an array");
  }

  const documentObj = xmlArray.find((item) => item["w:document"]);
  if (!documentObj) {
    throw new Error("No <w:document> tag found in the XML");
  }

  if (!Array.isArray(documentObj["w:document"])) {
    throw new Error("<w:document> tag is not an array");
  }

  const bodyObj = documentObj["w:document"].find((item) => item["w:body"]);
  if (!bodyObj) {
    throw new Error("No <w:body> tag found in the XML");
  }

  if (!Array.isArray(bodyObj["w:body"])) {
    throw new Error("<w:body> tag is not an array");
  }
  const bodyContent = bodyObj["w:body"].filter((item) => {
    return !item["w:sectPr"];
  });

  return bodyContent;
};

export const mergeDocx = (
  sourcePath: string,
  contentPath: string,
  {
    outputPath,
    pattern,
    insertStart = false,
    insertEnd = false,
  }: {
    outputPath?: string;
    pattern?: string;
    insertStart?: boolean;
    insertEnd?: boolean;
  }
) => {
  if (!sourcePath || !contentPath) {
    throw new Error("Missing source or content path");
  }
  if (
    !sourcePath.endsWith(".docx") ||
    !contentPath.endsWith(".docx") ||
    (outputPath && !outputPath.endsWith(".docx"))
  ) {
    throw new Error("Invalid file extension. Only .docx files are supported");
  }

  if (!pattern && !insertStart && !insertEnd) {
    throw new Error("Missing insert position or pattern");
  }

  if (!fs.existsSync(sourcePath)) {
    throw new Error(`Source file does not exist: ${sourcePath}`);
  }
  if (!fs.existsSync(contentPath)) {
    throw new Error(`Content file does not exist: ${contentPath}`);
  }

  const zip = new AdmZip(sourcePath);

  const templateXML = zip.readAsText(documentXMLPath, "utf8");
  if (!templateXML) {
    throw new Error(`Entry ${documentXMLPath} not found in ${sourcePath}`);
  }

  const contentZip = new AdmZip(contentPath);

  const contentXML = contentZip.readAsText(documentXMLPath, "utf8");
  if (!contentXML) {
    throw new Error(`Entry ${documentXMLPath} not found in ${contentPath}`);
  }

  const templateParsedXML = parseXML(templateXML);
  const contentParsedXML = parseXML(contentXML);
  const contentParsedXMLBody = getXmlBody(contentParsedXML);

  const finalParsedXML = mergeParsedXML(
    templateParsedXML,
    contentParsedXMLBody,
    { pattern, insertStart, insertEnd }
  );
  const finalXML = buildXML(finalParsedXML);

  // not using zip.updateFile(documentXMLPath, Buffer.from(finalXML, "utf8")) because then the docx file is corrupted
  const newZip = new AdmZip();
  zip.getEntries().forEach(function (zipEntry) {
    const entryBuffer =
      zipEntry.entryName == documentXMLPath
        ? Buffer.from(finalXML, "utf8")
        : zipEntry.getData();
    newZip.addFile(zipEntry.entryName, entryBuffer);
  });

  if (!outputPath) {
    return newZip.toBuffer();
  }

  return newZip.writeZip(outputPath);
};
