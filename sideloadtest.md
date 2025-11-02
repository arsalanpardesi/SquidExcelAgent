## Step 1: Start the Development Server Only

Instead of npm start, we will use a different command that only does the first job: starting the server.

    In your terminal, inside the OctalytiqExcelAgent folder, stop any running process with Ctrl+C.

    Run the following command:
    Bash

    npm run dev-server

    The terminal will show output like "Starting the dev server..." and "Compiled successfully." It will then wait. Keep this terminal window open. Your add-in's code is now being served from https://localhost:3000.

## Step 2: Sideload the Add-in Manually in Excel on the Web

Now we will act as the butler and tell Excel where to find your running code.

    Open your web browser and go to office.com. Log in with your new Microsoft 365 Business Basic account.

    Open Excel on the Web.

    In the Excel ribbon, go to the Insert tab and click Add-ins.

    In the Office Add-ins dialog that appears, click on "My Add-ins".

    At the bottom of the dialog, find the "Upload My Add-in" link and click it.

    A file browser will open. Navigate to your project folder (OctalytiqExcelAgent) and select the manifest.xml file. Click "Upload."

## Step 3: Test the Task Pane

After uploading the manifest, two things will happen:

    The add-in's task pane will immediately appear on the right side of the screen.

    A new button, "Show Taskpane," will be added to your Home ribbon for easy access later.

Click on some cells and press the "Run" button in the task pane. The cells should turn yellow.

Success! You now have a fully working development environment on Linux.

Your New Workflow

From now on, your development process will be:

    Run npm run dev-server in your terminal.

    Open your workbook in Excel on the Web. The add-in will already be there from the last time you uploaded the manifest.

    When you save changes to your code, the server will automatically recompile, and you can just refresh the browser tab to see your updates.

## Production notes
Step 1: Prepare the Manifest

Before you do anything else, run this command once in your terminal:
Bash

npm run manifest:dev

This command takes your manifest.xml template and replaces all the ~remoteAppUrl~ tokens with the correct https://localhost:3000 URL. Your manifest.xml file in your project folder is now valid and ready to be uploaded.