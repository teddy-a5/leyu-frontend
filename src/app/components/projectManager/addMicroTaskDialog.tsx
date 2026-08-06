import React, { useState, useEffect } from "react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialogLeft";
import { Button } from "@/components/ui/button";
import { Plus, Upload, FileAudio, Image } from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { projectTaskasAll, projectTaskasRelated } from "@/lib/hooks/useProject";
import {
  TaskResponse,
  TaskResponseData,
  ProjectTask,
} from "@/app/types/project";
import { toast } from "sonner";

interface AddMicroTaskDialogProps {
  tasks: TaskResponse[];
  taskMetadata: TaskResponseData;
  onSubmitSingle: (formData: {
    instruction: string;
    text: string;
    taskId: string;
    is_test: boolean;
    audioFiles?: File[];
    imageFiles?: File[];
  }) => void;
  onSubmitCsv: (uploadData: { file: File }) => void;
  onSubmitTask: (formData: {
    taskId: string;
    source_task_id: string;
    from_micro_task: boolean;
    from_data_set: boolean;
    limit: number | null;
  }) => void;
  onSubmitAudio: (uploadData: {
    files: File[];
    is_test: boolean;
    instruction: string;
  }) => void;
  onSubmitImage: (uploadData: {
    files: File[];
    is_test: boolean;
    instruction: string;
  }) => void;
  open: boolean;
  setOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

interface CsvRow {
  name: string;
  text: string;
  taskId: string;
  [key: string]: string;
  category:string;
  intent:string
}

const AddMicroTaskDialog: React.FC<AddMicroTaskDialogProps> = ({
  tasks,
  taskMetadata,
  onSubmitSingle,
  onSubmitCsv,
  onSubmitTask,
  onSubmitAudio,
  onSubmitImage,
  open,
  setOpen,
}) => {
  
  const [activeTab, setActiveTab] = useState<
    "Single" | "CsvUpload" | "ImportTask"
  >("Single");
  const [singleFormData, setSingleFormData] = useState({
    instruction: "",
    text: "",
    taskId: "",
    is_test: false,
  });

  const { data: projectTaskData } = projectTaskasRelated({
    task_id: taskMetadata.id,
  });

  const [audioFiles, setAudioFiles] = useState<File[]>([]);
  const [imageFiles, setImageFiles] = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [csvFile, setCsvFile] = useState<File | null>(null);
  const [csvData, setCsvData] = useState<CsvRow[]>([]);
  const [csvError, setCsvError] = useState<string | null>(null);
  const [source_task_id, setsource_task_id] = useState("");
  const [limit, setLimit] = useState("");
  const [from_micro_task, setfrom_micro_task] = useState(true);
  const [from_data_set, setfrom_data_set] = useState(false);

  // File size limits
  const MAX_CSV_SIZE = 7 * 1024 * 1024; // 7 MB
  const MAX_AUDIO_SIZE = 7 * 1024 * 1024; // 70MB
  const MAX_IMAGE_SIZE = 3 * 1024 * 1024; // 10MB

  const handleSwitchChange = (type: "microTask" | "dataSet") => {
    if (type === "microTask") {
      setfrom_micro_task(true);
      setfrom_data_set(false);
    } else {
      setfrom_micro_task(false);
      setfrom_data_set(true);
    }
  };

  // Reusable file size validator
  const validateFileSize = (file: File, type: "csv" | "audio" | "image"): boolean => {
    const maxSize = type === "audio" ? MAX_AUDIO_SIZE : type === "image" ? MAX_IMAGE_SIZE : MAX_CSV_SIZE;
    const maxSizeMB = type === "audio" ? 50 : type === "image" ? 10 : 10;

    if (file.size > maxSize) {
      toast.error(
        `${type === "csv" ? "CSV/XLSX" : type === "audio" ? "Audio" : "Image"} file is too large. Maximum allowed: ${maxSizeMB}MB`
      );
      return false;
    }
    return true;
  };

  // Parse CSV or XLSX file
  useEffect(() => {
    if (!csvFile) {
      setCsvData([]);
      setCsvError(null);
      return;
    }

    const fileExtension = csvFile.name.split(".").pop()?.toLowerCase();

    if (fileExtension === "csv") {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target?.result as string;
        Papa.parse<CsvRow>(text, {
          header: true,
          skipEmptyLines: true,
          complete: (result) => {
            const data = result.data;
            const requiredColumns: string[] = ["no", "text"];
            const allColumns:string[]=["name","text","taskId","category","intent"];
            const rawFields: string[] = result.meta.fields || [];
            const headers: string[] = rawFields.map((h) =>
              h.replace(/^\uFEFF/, "").trim().toLowerCase()
            );

            const hasText = headers.includes("text") || headers.includes("content");
            if (!hasText) {
              setCsvError("Missing required column: text");
              setCsvData([]);
              return;
            }

            const validRows: CsvRow[] = data.filter((row: any) => {
              const keys = Object.keys(row);
              const textKey = keys.find(
                (k) => k.replace(/^\uFEFF/, "").trim().toLowerCase() === "text" ||
                       k.replace(/^\uFEFF/, "").trim().toLowerCase() === "content"
              );
              return textKey && row[textKey]?.toString().trim();
            });

            if (validRows.length === 0) {
              setCsvError("No valid rows found in the file");
              setCsvData([]);
              return;
            }

            setCsvData(validRows);
            setCsvError(null);
          },
        });
      };
      reader.onerror = () => {
        setCsvError("Failed to read CSV file");
        setCsvData([]);
      };
      reader.readAsText(csvFile);
    } else if (fileExtension === "xlsx") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData: unknown[] = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
          });

          const rawHeaders = (jsonData[0] as string[]) || [];
          const headers = rawHeaders.map((h) =>
            String(h).replace(/^\uFEFF/, "").trim().toLowerCase()
          );

          const hasText = headers.includes("text") || headers.includes("content");
          if (!hasText) {
            setCsvError("Missing required column: text");
            setCsvData([]);
            return;
          }

          const rows: CsvRow[] = jsonData
            .slice(1)
            .reduce((acc: CsvRow[], row: unknown) => {
              if (Array.isArray(row)) {
                const rowData: CsvRow = {} as CsvRow;
                rawHeaders.forEach((header, index) => {
                  rowData[header] = row[index]?.toString() || "";
                });
                if (
                  row.some(
                    (val) => val !== undefined && val !== null && String(val).trim() !== ""
                  )
                ) {
                  acc.push(rowData);
                }
              }
              return acc;
            }, []);

          if (rows.length === 0) {
            setCsvError("No valid rows found in the XLSX file");
            setCsvData([]);
            return;
          }

          setCsvData(rows);
          setCsvError(null);
        } catch (error) {
          setCsvError(`Error parsing XLSX: ${(error as Error).message}`);
          setCsvData([]);
        }
      };
      reader.onerror = () => setCsvError("Failed to read XLSX file");
      reader.readAsArrayBuffer(csvFile);
    } else {
      setCsvError("Please upload a valid .csv or .xlsx file");
      setCsvData([]);
    }
  }, [csvFile]);

  // Reset CSV state when tab changes or dialog closes
  useEffect(() => {
    if (!open || activeTab !== "CsvUpload") {
      setCsvFile(null);
      setCsvData([]);
      setCsvError(null);
    }
  }, [open, activeTab]);

  const handleSingleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!singleFormData.text && audioFiles.length === 0 && imageFiles.length === 0) {
      toast.error("Please provide text or upload audio/image files");
      return;
    }
    onSubmitSingle({ 
      ...singleFormData, 
      audioFiles: audioFiles.length > 0 ? audioFiles : undefined,
      imageFiles: imageFiles.length > 0 ? imageFiles : undefined
    });
    setSingleFormData({
      instruction: "",
      text: "",
      taskId: "",
      is_test: false,
    });
    setAudioFiles([]);
    setImageFiles([]);
    setOpen(false);
  };

  const handleCsvSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (csvFile && csvData.length > 0 && !csvError) {
      onSubmitCsv({ file: csvFile });
      setCsvFile(null);
      setCsvData([]);
      setCsvError(null);
      setOpen(false);
    } else {
      toast.error("Please upload a valid CSV/XLSX file with data");
    }
  };

  const handleImageSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (imageFiles.length === 0) {
      toast.error("Please upload image files");
      return;
    }
    console.log('image')
    onSubmitImage({
      files: imageFiles,
      is_test: singleFormData.is_test,
      instruction: singleFormData.instruction,
    });
    setImageFiles([]);
    setSingleFormData({ ...singleFormData, instruction: "" });
    setOpen(false);
  };

  const handleAudioSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (audioFiles.length === 0) {
      toast.error("Please upload audio files");
      return;
    }
    onSubmitAudio({
      files: audioFiles,
      is_test: singleFormData.is_test,
      instruction: singleFormData.instruction,
    });
    setAudioFiles([]);
    setSingleFormData({ ...singleFormData, instruction: "" });
    setOpen(false);
  };

  const handleImportSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!source_task_id) {
      toast.error("Please select a task to import from");
      return;
    }
    onSubmitTask({
      taskId: taskMetadata.id,
      source_task_id,
      from_micro_task,
      from_data_set,
      limit: limit ? parseInt(limit) : null,
    });
    setsource_task_id("");
    setLimit("");
    setOpen(false);
  };

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger asChild>
        <Button className="bg-primary text-white hover:bg-blue-700 flex items-center gap-2">
          <Plus className="h-4 w-4" />
          New Micro Task
        </Button>
      </DialogTrigger>

      <DialogContent className="sm:max-w-[650px] h-[90vh] flex flex-col overflow-hidden">
        <DialogHeader>
          <p className="mb-4 font-bold text-lg">Add Micro Task </p>
        </DialogHeader>
        <DialogTitle />

        {/* Tabs */}
        <div className="border-b border-gray-200 mb-6 flex-shrink-0">
          <nav className="flex space-x-6">
            {["Single", 
              ...(taskMetadata?.taskType?.task_type !== "image-audio" && 
                  taskMetadata?.taskType?.task_type !== "image-text" ? ["CsvUpload"] : []), 
              "ImportTask"
            ].map((tab) => (
              <button
                key={tab}
                onClick={() => setActiveTab(tab as any)}
                className={`py-2 px-1 text-sm font-medium border-b-2 transition-colors ${
                  activeTab === tab
                    ? "border-primary text-primary"
                    : "border-transparent text-gray-500 hover:text-gray-700"
                }`}
              >
                {tab === "Single" && "Single Micro Task"}
                {tab === "CsvUpload" && "From CSV/XLSX"}
                {tab === "ImportTask" && "Import from Task"}
              </button>
            ))}
          </nav>
        </div>

        {/* Scrollable content area */}
        <div className="flex-1 overflow-y-auto pr-2">
          {/* === Single Micro Task === */}
          {activeTab === "Single" && (
            <form
              onSubmit={
                taskMetadata?.taskType?.task_type === "text-text" ||
                taskMetadata?.taskType?.task_type === "text-audio" 
                  ? handleSingleSubmit
                  : taskMetadata?.taskType?.task_type === "audio-text"
                  ? handleAudioSubmit
                  : taskMetadata?.taskType?.task_type === "image-audio" ||
                    taskMetadata?.taskType?.task_type === "image-text"
                  ? handleImageSubmit
                  : handleAudioSubmit
              }
              className="space-y-5"
            >
              {taskMetadata?.taskType?.task_type === "text-text" ||
              taskMetadata?.taskType?.task_type === "text-audio" ? (
                <>
                  <div>
                    <label className="block text-gray-700 font-medium mb-2">
                      Text *
                    </label>
                    <textarea
                      value={singleFormData.text}
                      onChange={(e) =>
                        setSingleFormData({
                          ...singleFormData,
                          text: e.target.value,
                        })
                      }
                      className="w-full p-3 border border-gray-300 rounded-lg focus:outline-none focus:border-blue-500 resize-vertical"
                      rows={5}
                      placeholder="Enter the text content..."
                      required
                    />
                  </div>
                </>
              ) : taskMetadata?.taskType?.task_type === "audio-text" ? (
                <div>
                  <label className="block text-gray-700 font-medium mb-2">
                    Audio File *
                  </label>
                  <input
                    type="file"
                    accept="audio/*"
                    multiple
                    onChange={(e) => {
                      const files = Array.from(e.target.files || []);
                      if (files.length > 5) {
                        toast.error("Maximum 5 audio files allowed");
                        e.target.value = "";
                        return;
                      }
                      if (files.length > 0) {
                        const validFiles = files.filter(file => validateFileSize(file, "audio"));
                        if (validFiles.length > 0) {
                          setAudioFiles(validFiles);
                        } else {
                          setAudioFiles([]);
                          e.target.value = "";
                        }
                      } else {
                        setAudioFiles([]);
                        e.target.value = "";
                      }
                    }}
                    className="w-full p-3 border border-gray-300 rounded-lg file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:bg-primary/10 file:text-primary hover:file:bg-primary/20"
                    required
                  />
                  {audioFiles.length > 0 && (
                    <div className="text-sm text-green-600 mt-2">
                      <p className="font-medium">Selected {audioFiles.length}/5 file(s):</p>
                      <ul className="list-disc list-inside ml-2 max-h-32 overflow-y-auto">
                        {audioFiles.map((file, index) => (
                          <li key={index} className="flex items-center justify-between">
                            <span>
                              <strong>{file.name}</strong> ({(file.size / 1024 / 1024).toFixed(2)} MB)
                            </span>
                            <button
                              type="button"
                              onClick={() => {
                                const newFiles = audioFiles.filter((_, i) => i !== index);
                                setAudioFiles(newFiles);
                              }}
                              className="ml-2 text-red-500 hover:text-red-700 text-xs"
                            >
                              Remove
                            </button>
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}
                  <p className="text-xs text-gray-500 mt-1">
                    Max size: 50MB per file • Max files: 5 • Formats: MP3, WAV, etc.
                  </p>
                </div>
              ) : taskMetadata?.taskType?.task_type === "image-audio" ||
                taskMetadata?.taskType?.task_type === "image-text" ? (
                <>
                  {/* Media Upload with Drag and Drop */}
                  <div>
                    <label className="block text-gray-700 font-medium mb-2">
                      Media
                    </label>
                    
                    {/* Drag and Drop Area */}
                    <div
                      onDragOver={(e) => {
                        e.preventDefault();
                        setIsDragging(true);
                      }}
                      onDragLeave={(e) => {
                        e.preventDefault();
                        setIsDragging(false);
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        setIsDragging(false);
                        const files = Array.from(e.dataTransfer.files);
                        if (files.length > 0) {
                          const file = files[0];
                          if (file.type.startsWith('image/')) {
                            if (validateFileSize(file, "image")) {
                              setImageFiles([file]);
                            }
                          } else {
                            toast.error("Please upload an image file");
                          }
                        }
                      }}
                      className={`border-2 border-dashed rounded-lg p-12 text-center transition-colors ${
                        isDragging
                          ? "border-primary bg-primary/5"
                          : "border-gray-300 bg-gray-50"
                      }`}
                    >
                      {imageFiles.length === 0 ? (
                        <>
                          <div className="flex justify-center mb-4">
                            <Upload className="h-12 w-12 text-gray-400" />
                          </div>
                          <p className="text-gray-600 font-medium mb-2">
                            Drag and Drop a file
                          </p>
                          <p className="text-sm text-gray-500 mb-4">
                            Minimum 1MB and Max 10 MB each can be uploaded
                          </p>
                          <p className="text-sm text-gray-500 mb-4">or</p>
                          <label className="inline-block">
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) => {
                                const files = Array.from(e.target.files || []);
                                if (files.length > 0) {
                                  const file = files[0];
                                  if (validateFileSize(file, "image")) {
                                    setImageFiles([file]);
                                  } else {
                                    e.target.value = "";
                                  }
                                }
                              }}
                              className="hidden"
                              id="image-file-input"
                            />
                            <span className="px-6 py-2 border border-gray-300 rounded-lg text-gray-600 hover:bg-gray-100 cursor-pointer inline-block">
                              Browse Files
                            </span>
                          </label>
                        </>
                      ) : (
                        <div className="space-y-4">
                          <div className="flex items-center justify-center">
                            <div className="relative">
                              <img
                                src={URL.createObjectURL(imageFiles[0])}
                                alt="Preview"
                                className="w-48 h-48 object-cover rounded-lg border-2 border-gray-200"
                              />
                              <button
                                type="button"
                                onClick={() => {
                                  setImageFiles([]);
                                  const fileInput = document.getElementById('image-file-input') as HTMLInputElement;
                                  if (fileInput) fileInput.value = "";
                                }}
                                className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-8 h-8 flex items-center justify-center hover:bg-red-600 shadow-lg"
                              >
                                ×
                              </button>
                            </div>
                          </div>
                          <div className="text-center">
                            <p className="text-sm font-medium text-gray-700">
                              {imageFiles[0].name}
                            </p>
                            <p className="text-xs text-gray-500">
                              {(imageFiles[0].size / 1024 / 1024).toFixed(2)} MB
                            </p>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </>
              ) : null}

              <div>
                <label className="block text-gray-700 font-medium mb-2">
                  Instruction
                </label>
                <textarea
                  value={singleFormData.instruction}
                  onChange={(e) =>
                    setSingleFormData({
                      ...singleFormData,
                      instruction: e.target.value,
                    })
                  }
                  className="w-full p-3 border border-gray-300 rounded-lg focus:outline-none focus:border-blue-500 resize-vertical"
                  rows={4}
                  placeholder="Optional instructions for contributors..."
                />
              </div>

              {taskMetadata?.require_contributor_test && (
                <label className="flex items-center gap-3 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={singleFormData.is_test}
                    onChange={(e) =>
                      setSingleFormData({
                        ...singleFormData,
                        is_test: e.target.checked,
                      })
                    }
                    className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                  />
                  <span className="text-gray-700">Mark as test task</span>
                </label>
              )}

              <div className="flex justify-end gap-3 pt-4">
                <Button
                  variant="outline"
                  type="button"
                  onClick={() => setOpen(false)}
                >
                  Cancel
                </Button>
                <Button
                  type="submit"
                  className="bg-primary hover:bg-primary/90"
                >
                  Create Micro Task
                </Button>
              </div>
            </form>
          )}

          {/* === CSV/XLSX Upload === */}
          {activeTab === "CsvUpload" &&
            (taskMetadata?.taskType?.task_type === "text-text" ||
              taskMetadata?.taskType?.task_type === "text-audio" ||
              taskMetadata?.taskType?.task_type === "audio-text" ||
              taskMetadata?.taskType?.task_type === "image-text" ||
              taskMetadata?.taskType?.task_type === "image-audio") && (
              <form onSubmit={handleCsvSubmit} className="space-y-5">
                <div>
                  <label className="block text-gray-700 font-medium mb-2">
                    Upload CSV or XLSX File *
                  </label>
                  <input
                    type="file"
                    accept=".csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    multiple={false}
                    onChange={(e) => {
                      const file = e.target.files?.[0] || null;
                      if (file && validateFileSize(file, "csv")) {
                        setCsvFile(file);
                        setCsvError(null);
                      } else {
                        setCsvFile(null);
                        setCsvData([]);
                        setCsvError(null);
                        e.target.value = "";
                      }
                    }}
                    className="w-full p-3 border border-gray-300 rounded-lg file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
                  />
                  {csvFile && (
                    <p className="text-sm text-green-600 mt-2">
                      Selected: <strong>{csvFile.name}</strong> (
                      {(csvFile.size / 1024 / 1024).toFixed(2)} MB)
                    </p>
                  )}
                  <p className="text-xs text-gray-500 mt-1">
                    Max size: 2MB • Required columns:{" "}
                    <code className="bg-gray-100 px-1 rounded">no</code>,{" "}
                    <code className="bg-gray-100 px-1 rounded">text</code>
                  </p>
                  {csvError && (
                    <p className="text-sm text-red-600 mt-2">{csvError}</p>
                  )}
                </div>

                {csvData.length > 0 && (
                  <div className="bg-gray-50 p-4 rounded-lg">
                    <h4 className="font-semibold text-gray-800 mb-3">
                      Preview ({csvData.length} rows)
                    </h4>
                    <div className="overflow-x-auto max-h-64 border rounded">
                      <table className="min-w-full text-sm">
                        <thead className="bg-gray-100 sticky top-0">
                          <tr>
                            <th className="px-3 py-2 text-left">No</th>
                            <th className="px-3 py-2 text-left">Text</th>
                            <th className="px-3 py-2 text-left">Category</th>
                            <th className="px-3 py-2 text-left">Intent</th>
                          </tr>
                        </thead>
                        <tbody>
                          {csvData.slice(0, 10).map((row, i) => (
                            <tr key={i} className="border-t">
                              <td className="px-3 py-2">
                                {row.no || row.name}
                              </td>
                              <td className="px-3 py-2 max-w-md truncate">
                                {row.text}
                              </td>
                              <td className="px-3 py-2 max-w-md truncate">
                                {row.category}
                              </td>
                              <td className="px-3 py-2 max-w-md truncate">
                                {row.intent}
                              </td>
                            </tr>
                          ))}
                          {csvData.length > 10 && (
                            <tr>
                              <td
                                colSpan={2}
                                className="px-3 py-2 text-center text-gray-500 italic"
                              >
                                ... and {csvData.length - 10} more rows
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                <div className="flex justify-end gap-3 pt-4">
                  <Button
                    variant="outline"
                    type="button"
                    onClick={() => setOpen(false)}
                  >
                    Cancel
                  </Button>
                  <Button
                    type="submit"
                    disabled={!csvFile || csvData.length === 0 || !!csvError}
                    className="bg-primary hover:bg-primary/90 disabled:opacity-50"
                  >
                    Upload and Create
                  </Button>
                </div>
              </form>
            )}

          {/* === Import from Task === */}
          {activeTab === "ImportTask" && (
            <form onSubmit={handleImportSubmit} className="space-y-5">
              <div>
                <label className="block text-gray-700 font-medium mb-2">
                  Select Task to Import From
                </label>
                <select
                  value={source_task_id}
                  onChange={(e) => setsource_task_id(e.target.value)}
                  className="w-full p-3 border border-gray-300 rounded-lg focus:outline-none focus:border-blue-500"
                  required
                >
                  <option value="">Choose a task...</option>
                  {(Array.isArray(projectTaskData?.data)
                    ? (projectTaskData?.data ?? [])
                    : []
                  ).map((task: any) => (
                    <option key={task.id} value={task.id}>
                      {task.name} ({task.taskType?.task_type})
                    </option>
                  ))}
                </select>
              </div>

              <div className="flex gap-6">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={from_micro_task}
                    onChange={() => handleSwitchChange("microTask")}
                    className="text-blue-600"
                  />
                  <span>From MicroTasks</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={from_data_set}
                    onChange={() => handleSwitchChange("dataSet")}
                    className="text-blue-600"
                  />
                  <span>From Dataset</span>
                </label>
              </div>

              <div>
                <label className="block text-gray-700 font-medium mb-2">
                  Limit (optional)
                </label>
                <input
                  type="number"
                  value={limit}
                  onChange={(e) => setLimit(e.target.value)}
                  min="1"
                  placeholder="e.g., 100"
                  className="w-full p-3 border border-gray-300 rounded-lg focus:outline-none focus:border-blue-500"
                />
                <p className="text-xs text-gray-500 mt-1">
                  Leave empty to import all
                </p>
              </div>

              <div className="flex justify-end gap-3 pt-4">
                <Button
                  variant="outline"
                  type="button"
                  onClick={() => setOpen(false)}
                >
                  Cancel
                </Button>
                <Button
                  type="submit"
                  disabled={!source_task_id}
                  className="bg-primary hover:bg-primary/90"
                >
                  Import Microtasks
                </Button>
              </div>
            </form>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
};

export default AddMicroTaskDialog;
