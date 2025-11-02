# Octalytiq - Squid Excel AI Agent

This is an open-source, fully agentic AI copilot for Microsoft Excel. The current version supports Google's Gemini API. It uses LangGraph for orchestrating the agentic workflow. The workflow is built to understand complex, multi-step tasks from a natural language prompt, read your data, and execute a plan directly in your worksheet.

This current version is comparable to other Excel copilot agents available commercially, but this is for free other than the Gemini API credits 😄.

## Demo 1

[![Watch the Demo 1](https://img.youtube.com/vi/m04iPe7FuBU/hqdefault.jpg)](https://www.youtube.com/watch?v=m04iPe7FuBU)

## Demo 2

[![Watch the Demo 2](https://img.youtube.com/vi/2ajH7caXuDc/hqdefault.jpg)](https://www.youtube.com/watch?v=2ajH7caXuDc)

## Features

    🤖 Fully Agentic: Uses a LangGraph-based agent to route between tasks, break down complex goals into sub-tasks, and generate multi-step plans.

    📄 PDF Analysis: Can read and incorporate data from an uploaded PDF file into its reasoning and planning process.

    🧠 Powered by AI: Natively configured to use Google's powerful Gemini family of models for all reasoning and plan generation.

    ⚡ Streaming Responses: Streams its thoughts, plans, and status updates to the taskpane in real-time, just like a modern AI chatbot.

## Tech Stack

    Frontend: Microsoft Excel Office Add-in (TypeScript, HTML, CSS)

    Backend: Node.js, Express

    AI Engine: LangGraph.js, Google Gemini

## 🚀 Getting Started

This project is in two parts: the server (the AI backend) and the main root folder (the Excel add-in frontend). You will need two terminals running at the same time.

## 1. Clone & Install

First, clone the repository and install all dependencies for both the frontend and backend.
Bash

### 1. Clone the project
git clone https://github.com/arsalanpardesi/SquidExcelAgent.git

### 2. Go into the project root
cd your-repo-folder

### 3. Install frontend dependencies
npm install

### 4. Go into the server folder
cd server

### 5. Install backend dependencies
npm install

## 2. Configure the Backend (Server)

You must provide a Google Gemini API key for the agent to work.

    In the /server folder, copy the .env.example file and rename it to .env.

    Open your new .env file with a text editor.

    Paste your Google Gemini API key into the GEMINI_API_KEY variable:
    Code snippet

    GEMINI_API_KEY="YOUR_API_KEY_GOES_HERE"

## 3. Run the Project

Open two terminals for the next two steps.

Terminal 1: Start the Backend Server

Bash

### Navigate to the /server folder
cd server

### Start the backend (it will run on http://localhost:3001)
npm run dev

Terminal 2: Start the Frontend Add-in

Bash

### Navigate to the main project root folder
cd .. 

### Start the frontend (it will run on https://localhost:3000)
npm run dev-server

⚙️ How to Load the Add-in in Excel

You must "sideload" the add-in's manifest file to use it in Excel.

## 5. Sideload the Add-in in Excel

    Go to Excel on the Web (this is the most reliable way).

    Open a blank workbook.

    On the Home tab, click Add-ins (it may be under the ... menu).

    In the "Office Add-ins" window, select the "My Add-ins" tab.

    At the bottom of the window, click "Upload My Add-in".

    A dialog will open. Click "Browse..." and select the manifest.xml file located in the root of your project folder.

    Click "Upload".

    The add-in is now installed! You should see a new button for your agent on the Home ribbon.

## 6. Fix the Security Certificate (First Time Only)

The first time you run this, Excel will show an error because the dev server uses a "self-signed" certificate. You must tell your browser to trust it.

    Open your web browser (Chrome, Edge, etc.).

    In the address bar, navigate to https://localhost:3000

    You will see a large security warning page (e.g., "Your connection is not private").

    Click the "Advanced" button.

    Click the link that says "Proceed to localhost (unsafe)".

    A blank page or an error will load. This is fine. You can now close the browser tab.

    Go back to Excel and click your add-in's button on the ribbon. It will now load correctly.

### 🧪 Testing & Compatibility

This add-in has been primarily tested on Excel on the Web with the server running on linux.

It should be compatible with the modern Excel desktop client on Windows and Mac, but it has not been extensively tested in those environments.

### ⚖️ License

This project is licensed under the Apache 2.0 License. See the LICENSE file for details.