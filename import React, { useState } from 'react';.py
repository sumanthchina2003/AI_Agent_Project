import React, { useState, useCallback } from 'react';
import { Upload, Search, Database, Download, RefreshCw } from 'lucide-react';
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Alert,
  AlertDescription,
  AlertTitle,
} from "@/components/ui/alert";

const API_URL = 'http://localhost:5000/api';

const Dashboard = () => {
  // State management
  const [dataSource, setDataSource] = useState('file');
  const [fileData, setFileData] = useState(null);
  const [selectedColumn, setSelectedColumn] = useState('');
  const [columns, setColumns] = useState([]);
  const [prompt, setPrompt] = useState('');
  const [results, setResults] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState('');
  const [preview, setPreview] = useState([]);

  // Handle file upload
  const handleFileUpload = useCallback(async (event) => {
    const file = event.target.files[0];
    if (file) {
      try {
        const reader = new FileReader();
        reader.onload = async (e) => {
          const text = e.target.result;
          const rows = text.split('\n');
          const headers = rows[0].split(',');
          
          setColumns(headers);
          setPreview(
            rows.slice(1, 6).map(row => {
              const values = row.split(',');
              return headers.reduce((obj, header, index) => {
                obj[header] = values[index];
                return obj;
              }, {});
            })
          );
        };
        reader.readAsText(file);
      } catch (err) {
        setError('Error processing file: ' + err.message);
      }
    }
  }, []);

  // Handle Google Sheets connection
  const handleSheetsConnection = useCallback(async (sheetsUrl) => {
    try {
      setIsProcessing(true);
      const response = await fetch(`${API_URL}/sheets/connect`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ spreadsheetId: sheetsUrl }),
      });
      
      const data = await response.json();
      
      if (data.success) {
        setColumns(data.data[0]);
        setPreview(data.data.slice(1, 6));
      } else {
        setError(data.error);
      }
    } catch (err) {
      setError('Error connecting to Google Sheets: ' + err.message);
    } finally {
      setIsProcessing(false);
    }
  }, []);

  // Process data and extract information
  const handleProcessing = useCallback(async () => {
    try {
      setIsProcessing(true);
      setError('');
      
      const entities = preview.map(row => row[selectedColumn]);
      
      const response = await fetch(`${API_URL}/process`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          sourceType: dataSource,
          mainColumn: selectedColumn,
          promptTemplate: prompt,
          entities,
        }),
      });
      
      const data = await response.json();
      
      if (data.success) {
        setResults(data.results);
      } else {
        setError(data.error);
      }
    } catch (err) {
      setError('Error processing data: ' + err.message);
    } finally {
      setIsProcessing(false);
    }
  }, [dataSource, selectedColumn, prompt, preview]);

  // Download results as CSV
  const handleDownload = useCallback(() => {
    try {
      const csvContent = [
        ['Entity', 'Extracted Information'],
        ...results.map(result => [result.entity, result.extracted_info])
      ].map(row => row.join(',')).join('\n');

      const blob = new Blob([csvContent], { type: 'text/csv' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.setAttribute('hidden', '');
      a.setAttribute('href', url);
      a.setAttribute('download', 'results.csv');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err) {
      setError('Error downloading results: ' + err.message);
    }
  }, [results]);

  return (
    <div className="container mx-auto p-4 space-y-6">
      {/* Data Source Selection */}
      <Card>
        <CardHeader>
          <CardTitle>Data Source</CardTitle>
          <CardDescription>Choose your data input method</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="flex space-x-4">
            <Button
              variant={dataSource === 'file' ? 'default' : 'outline'}
              onClick={() => setDataSource('file')}
            >
              <Upload className="mr-2 h-4 w-4" />
              CSV Upload
            </Button>
            <Button
              variant={dataSource === 'sheets' ? 'default' : 'outline'}
              onClick={() => setDataSource('sheets')}
            >
              <Database className="mr-2 h-4 w-4" />
              Google Sheets
            </Button>
          </div>

          {dataSource === 'file' ? (
            <Input
              type="file"
              accept=".csv"
              onChange={handleFileUpload}
              className="w-full"
            />
          ) : (
            <div className="space-y-4">
              <Input
                type="text"
                placeholder="Enter Google Sheets URL"
                className="w-full"
                onBlur={(e) => handleSheetsConnection(e.target.value)}
              />
            </div>
          )}
        </CardContent>
      </Card>

      {/* Preview Data */}
      {preview.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle>Data Preview</CardTitle>
          </CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  {columns.map((column) => (
                    <TableHead key={column}>{column}</TableHead>
                  ))}
                </TableRow>
              </TableHeader>
              <TableBody>
                {preview.map((row, index) => (
                  <TableRow key={index}>
                    {columns.map((column) => (
                      <TableCell key={column}>{row[column]}</TableCell>
                    ))}
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}

      {/* Query Configuration */}
      {columns.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle>Query Configuration</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="space-y-2">
              <Label>Select Main Column</Label>
              <Select onValueChange={setSelectedColumn}>
                <SelectTrigger>
                  <SelectValue placeholder="Select a column" />
                </SelectTrigger>
                <SelectContent>
                  {columns.map((col) => (
                    <SelectItem key={col} value={col}>
                      {col}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <Label>Query Prompt</Label>
              <Input
                placeholder="Get me the email address of {company}"
                value={prompt}
                onChange={(e) => setPrompt(e.target.value)}
              />
              <p className="text-sm text-gray-500">
                Use {'{company}'} as a placeholder for each entity
              </p>
            </div>

            <Button
              onClick={handleProcessing}
              disabled={!selectedColumn || !prompt || isProcessing}
              className="w-full"
            >
              {isProcessing ? (
                <>
                  <RefreshCw className="mr-2 h-4 w-4 animate-spin" />
                  Processing...
                </>
              ) : (
                <>
                  <Search className="mr-2 h-4 w-4" />
                  Start Processing
                </>
              )}
            </Button>
          </CardContent>
        </Card>
      )}

      {/* Results */}
      {results.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle>Results</CardTitle>
          </CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Entity</TableHead>
                  <TableHead>Extracted Information</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {results.map((result, index) => (
                  <TableRow key={index}>
                    <TableCell>{result.entity}</TableCell>
                    <TableCell>{result.extracted_info}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>

            <div className="mt-4