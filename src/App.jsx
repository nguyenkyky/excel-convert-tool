import React, { useRef, useState } from "react";
import PropTypes from "prop-types";
import { useDropzone } from "react-dropzone";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import JSZip from "jszip";
import "./style.scss";
import Button from "react-bootstrap/Button";
import { FormControl } from "react-bootstrap";
import axios from "axios";


const ExcelTool = () => {
    const [selectedFile, setSelectedFile] = useState(null);
    const [convertedTxtLinks, setConvertedTxtLinks] = useState([]);
    const [zipFileName, setZipFileName] = useState("Converted");
    const jsonDataRef = useRef([]);
    const fieldIndexesRef = useRef([]);

    const onDrop = (acceptedFiles) => {
        if (acceptedFiles && acceptedFiles.length > 0) {
            
            const file = acceptedFiles[0];
            if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
                setSelectedFile(acceptedFiles[0]);
                setConvertedTxtLinks([]); // Reset the converted txt links
                setZipFileName(getBaseFileName(acceptedFiles[0].name)); // Đặt tên mặc định file ZIP giống với file excel
            } else {
                alert("Chọn đúng định dạng file (.xlsx hoặc .xls)");
            }
        }
    };



    // Lấy tên file excel
    function getBaseFileName(fileName) {
        const lastDotIndex = fileName.lastIndexOf(".");
        if (lastDotIndex !== -1) {
            return fileName.substring(0, lastDotIndex);
        }
        return fileName;
    }
    const gradeLevels = [
        "<Preschool>",
        "<Kindergarten>",
        "<Grade 1>",
        "<Grade 2>",
        "<Grade 3>",
        "<Grade 4>",
        "<Grade 5>",
        "<Grade 6>",
        "<Grade 7>",
        "<Grade 8>",
        "<Grade 9>",
        "<Grade 10>",
        "<Grade 11>",
        "<Grade 12>",
    ];
    const newGradeLevels = [
        "Preschool",
        "Kindergarten",
        "Grade 1",
        "Grade 2",
        "Grade 3",
        "Grade 4",
        "Grade 5",
        "Grade 6",
        "Grade 7",
        "Grade 8",
        "Grade 9",
        "Grade 10",
        "Grade 11",
        "Grade 12",
    ];

    const convertExcelToTxt = () => {
        if (selectedFile) {
            const reader = new FileReader();
            reader.onload = (event) => {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: "array" });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                    header: 1,
                });
                jsonDataRef.current = XLSX.utils.sheet_to_json(firstSheet, {
                    header: 1,
                });

                // Lấy các fields ở hàng đầu tiên
                const headerRow = jsonData[0];

                const fields = [
                    "Title",
                    "Tags",
                    "Topics",
                    "CCSS",
                    "Categories",
                    "Grade",
                    "Description",
                ];

                // Lấy dữ liệu của Start grade và End grade
                const gradeFields = ["Start grade", "End grade"];
                const fieldIndexes = findFieldIndexes(headerRow, fields);
                fieldIndexesRef.current = findFieldIndexes(headerRow, fields);
                const gradeFieldIndexes = findFieldIndexes(
                    headerRow,
                    gradeFields
                );

                const txtFileLinks = [];

                jsonData.slice(1).forEach((rowData, rowIndex) => {
                    // Kiểm tra xem có dữ liệu trong dòng rowData hay không thì mới convert
                    const hasData = fields.some(
                        (field) =>
                            rowData[fieldIndexes[field]] !== undefined &&
                            rowData[fieldIndexes[field]] !== ""
                    );
                    if (hasData) {
                        const startGrade =
                            rowData[gradeFieldIndexes["Start grade"]];
                        const endGrade =
                            rowData[gradeFieldIndexes["End grade"]];
                        let gradeRange = [];
                        // console.log(startGrade, endGrade);

                        // Lấy khoảng Grade
                        if (startGrade && endGrade) {
                            let startIndex = gradeLevels.indexOf(startGrade);
                            let endIndex = gradeLevels.indexOf(endGrade);
                            if (startIndex === -1 || endIndex === -1) {
                                startIndex = newGradeLevels.indexOf(startGrade);
                                endIndex = newGradeLevels.indexOf(endGrade);
                                gradeRange = gradeLevels.slice(
                                    startIndex,
                                    endIndex + 1
                                );
                            } else {
                                gradeRange = gradeLevels.slice(
                                    startIndex,
                                    endIndex + 1
                                );
                                //   console.log(gradeRange);
                            }
                        }

                        const title = rowData[fieldIndexes["Title"]];
                        if(!title) return;

                        // Lấy dữ liệu trên từng hàng ứng với các trường trong mảng fields[]
                        const txtContentArray = fields.map((field) => {
                            const cellValue =
                                rowData[fieldIndexes[field]] || "";
                            if (field === "Tags" || field === "Topics" || field === "CCSS" || field === "Categories") {
                                if(cellValue) {

                                    const dataParts = cellValue
                                        .split(",")
                                        .map((item) => `<${item.trim()}>`);
                                    return `${field}:${dataParts.join(", ")}`;
                                }
                            }
                            if (field === "Grade") {
                                return `Grade:${gradeRange.join(", ")}`;
                            } else {
                                return `${field}:${
                                    cellValue ? `<${cellValue}>` : ""
                                }`;
                            }
                        });

                        const txtBlob = new Blob([txtContentArray.join("\n")], {
                            type: "text/plain",
                        });
                        const txtDownloadLink = URL.createObjectURL(txtBlob);
                        txtFileLinks.push({title, link: txtDownloadLink});
                    }
                });

                setConvertedTxtLinks(txtFileLinks);
            };
            reader.readAsArrayBuffer(selectedFile);
        }
    };

    // Đánh dấu vị trí của các phần tử trong fields(đánh dấu vị trí cột)
    const findFieldIndexes = (headerRow, fields) => {
        const fieldIndexes = {};

        headerRow.forEach((cellValue, columnIndex) => {
            if (fields.includes(cellValue)) {
                fieldIndexes[cellValue] = columnIndex;
            }
        });

        return fieldIndexes;
    };

    // Đoạn này warning
    const { getRootProps, getInputProps } = useDropzone({
        onDrop,
        accept: ".xlsx, .xls",
        maxFiles: 1,
    });

    const downloadZip = async () => {
        if (convertedTxtLinks.length > 0) {
            const zip = new JSZip();

            await Promise.all(
                convertedTxtLinks.map(({title,link}, index) => {
                    const formatTitle = title.replace(/[/\\?%*:|"<>]/g, "-"); // Thay thế các ký tự không hợp lệ khi đặt tên bằng dấu gạch ngang
                    // const title = jsonDataRef.current[index+1][fieldIndexesRef.current["Title"]];
                    // const fileName = title ? `${title}.txt` : `data_${index + 1}.txt`; // Đặt tên tệp txt
                    return fetch(link)
                        .then((response) => response.blob())
                        .then((blob) => {
                            zip.file(`${formatTitle}.txt`, blob);
                        });
                })
            );

            zip.generateAsync({ type: "blob" }).then((content) => {
                saveAs(content, `${zipFileName}.zip`);
            });
        }
    };

    return (
        <div className="convert-tool">
            <div {...getRootProps()} className="dropzone">
                <input {...getInputProps()} />
                <p>Kéo thả file Excel, hoặc Click để tải lên</p>
            </div>
            {selectedFile && (
                <div className="selected-file">
                    <p>File đã chọn: {selectedFile.name}</p>
                    <Button className="text-white" onClick={convertExcelToTxt}>
                        Chuyển đổi sang txt
                    </Button>
                </div>
            )}
            {convertedTxtLinks.length > 0 && (
                <div className="download">
                    <p style={{ fontSize: "13px" }}>
                        Converted {selectedFile.name}
                    </p>

                    <p>Nhập tên ZIP file</p>
                    <input
                        className="name-form"
                        type="text"
                        value={zipFileName}
                        onChange={(e) => setZipFileName(e.target.value)}
                    />

                    <p>Tải về ZIP file</p>
                    <Button className="text-white" onClick={downloadZip}>
                        Download
                    </Button>
                </div>
            )}
        </div>
    );
};

export default ExcelTool;
