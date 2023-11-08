"use client";

import createReport, { listCommands } from "docx-templates";
import { useState } from "react";

const downloadURL = (data, fileName) => {
  const a = document.createElement("a");
  a.href = data;
  a.download = fileName;
  document.body.appendChild(a);
  a.style = "display: none";
  a.click();
  a.remove();
};

const saveDataToFile = (data, fileName, mimeType) => {
  const blob = new Blob([data], { type: mimeType });
  const url = window.URL.createObjectURL(blob);
  downloadURL(url, fileName, mimeType);
  setTimeout(() => {
    window.URL.revokeObjectURL(url);
  }, 1000);
};

const initialData = {
  name: "John",
  phone: "123456",
  imageURL:
    "https://images.pexels.com/photos/268533/pexels-photo-268533.jpeg?auto=compress&cs=tinysrgb&w=1260&h=750&dpr=1",
  showTable: true,
  link: {
    url: "https://www.google.com/",
    label: "Link to Google",
  },
  html: {
    title: "HTML title",
    description: "HTML description ",
  },
  items: [
    { id: "1", title: "title1", description: "description1" },
    { id: "2", title: "title2", description: "description2" },
  ],
};
export default function Home() {
  const [file, setFile] = useState();
  const [data, setData] = useState(JSON.stringify(initialData));
  const readFileIntoArrayBuffer = (fd) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = reject;
      reader.onload = () => {
        resolve(reader.result);
      };
      reader.readAsArrayBuffer(fd);
    });
  };

  const generateFile = async () => {
    if(!file){
      alert("please Select docx file first");
      return;
    }
    const template = await readFileIntoArrayBuffer(file);
    // const commands = await listCommands(template, ["{", "}"]);
    // console.log({ commands });
    let templateData;
    try {
      templateData = JSON.parse(data);
    } catch (e) {}
    if (!templateData) {
      alert("please enter valid json string");
      return;
    }
    const report = await createReport({
      template,
      additionalJsContext: {
        renderImageFromURL: async (url, size = 3) => {
          const resp = await fetch(url);
          const buffer = resp.arrayBuffer
            ? await resp.arrayBuffer()
            : await resp.buffer();
          return { width: size, height: size, data: buffer, extension: ".png" };
        },
      },
      errorHandler: (err, command_code) => {
        return "command failed!";
      },
      data: templateData,
    });
    saveDataToFile(
      report,
      "report.docx",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
  };
  const onTemplateChosen = async (event) => {
    const file = event.target.files[0];
    setFile(file);
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        width: "700px",
        margin: "0 auto",
      }}
    >
      <h1>Docx Example</h1>
      <a href="/example.docx" download="example">
        Example File
      </a>
      <input
        type="file"
        onChange={onTemplateChosen}
        style={{ margin: "10px 0px" }}
      />
      <textarea
        value={data}
        onChange={(e) => setData(e.target.value)}
        rows={8}
      />
      <button
        type="button"
        onClick={generateFile}
        style={{ width: "200px", margin: "10px 0px" }}
      >
        Replace Template Data
      </button>
    </div>
  );
}
