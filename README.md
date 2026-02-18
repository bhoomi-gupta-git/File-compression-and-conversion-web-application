# File-compression-and-conversion-web-application

It's a Python Flask based web application which is used to compress files of different formats and also convert files into different formats as per user convenience.

## Introduction
The rapid growth of digital data has increased the size of images, videos, PDFs and office documents. This creates problems in storage management, file sharing and uploading. This project provides an offline file compression system which reduces file size while maintaining quality.

## Problem Statement
Most compression tools are cloud-based and require uploading files to external servers which can cause privacy issues, internet dependency and file upload restrictions. This project solves this problem by providing a completely offline compression system.

## Objectives
- To build a secure offline file compression system using Python and Flask
- To support compression for multiple file formats
- To provide a simple upload → compress → download interface
- To ensure data privacy by not uploading files to any external server
- To maintain file quality and structure after compression

## Features
- Offline file compression (no internet required)
- Multi-format file compression
- Simple and user-friendly interface
- Error handling and validation
- Output file download option
- Temporary file cleanup after compression

## Supported File Formats
- Images (JPG, JPEG, PNG)
- Videos (MP4, MOV, MKV)
- PDF Documents
- Word Documents (.docx)
- PowerPoint Presentations (.pptx)
- Excel Workbooks (.xlsx)
- ZIP Archives (.zip)

## Technology Stack
- Python
- Flask Framework
- HTML, CSS
- Jinja2 Template Engine

## Libraries Used
- Pillow (Image compression)
- MoviePy + FFmpeg (Video compression)
- Ghostscript (PDF compression)
- python-docx (Word compression)
- python-pptx (PowerPoint compression)
- openpyxl (Excel compression)
- ZipFile module (ZIP recompression)

## Workflow
1. User uploads a file
2. System validates the file format
3. File type is detected automatically
4. Compression function is applied based on file type
5. Compressed file is saved and provided for download
6. Temporary files are cleaned automatically

## Results
- JPEG images compressed by 40% to 70%
- PNG images compressed by 10% to 30%
- Videos compressed by 30% to 80%
- PDFs compressed by 30% to 60%
- Word documents compressed by 10% to 40%
- PowerPoint compressed by 20% to 50%
- Excel compressed by 5% to 20%
- ZIP recompression achieved 5% to 20% reduction

## Security and Privacy
This system is fully offline and does not upload files to the cloud. No external API calls are made. All processing happens on the user's local machine, ensuring complete privacy.

## Limitations
- Video compression is slow for large video files
- Office file compression depends on embedded media
- Batch compression is not included in this version

## Future Scope
- Batch compression feature
- Compression presets (low/medium/high)
- Faster video compression optimization
- Support for more file formats
- Adding file conversion module in major project version
