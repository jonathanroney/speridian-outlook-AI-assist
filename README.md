# speridian-outlook-AI-assist

An AI assist for Outlook (classic) created as a project for a Speridian Technologies internship. The tools created are:

- Generate Email
- Spell Check
- Reply Assist
- Language Conversion
- Chatbot

# Overview

## Objective

This add-in was created as an internship project that aimed to extensively implement basic AI features into the Outlook ecosystem - this helps users efficiently manage their communication in both casual and professional conversations.

## Programs / Technology stack

This project is a VSTO add-in created for Outlook (classic) and was created in C# in Visual Studio. It also made use of .NET and the Google Gemini AI API.

## File structure

<img width="288" height="470" alt="Picture1" src="https://github.com/user-attachments/assets/57ceb8eb-d9ff-4c80-a5d2-1700868b7767" />

# Architecture

### Ribbon
The add-in consists of a separate "AI assist" tab in the ribbon, where all 5 of the separate buttons to use the tools are located.

<img width="752" height="89" alt="Picture2" src="https://github.com/user-attachments/assets/5648e88d-84cc-4f91-b417-b7fbf72d0168" />


### ThisAddIn.cs

The file that controls the add-in's behaviour on startup and shutdown. Note: This remained largely unused except for a small fix concerning the thread usage of the chatbot – ensured that WindowsFormSynchronizationContext was installed to allow await functions to continue on the UI thread rather than a background thread.

### Forms

Each button in the ribbon opens the respective form for the tool, where the user is prompted to input the required information. Additionally, there is a loading form that is used to indicate when an action is in progress (such as an AI API request) – this defaults to the spell check tool's loading message.

<img width="392" height="121" alt="Picture3" src="https://github.com/user-attachments/assets/316932e4-73b8-407a-a366-8b1a21d9bf37" />

### Services  
Rather than handling the computation behind each tool in the forms, they call functions from the service classes which actually perform the tasks, including AI API calls, parsing, etc. The only tasks performed by the ribbon are related to opening and closing the various Outlook object models like mailItems, Inspectors, and Explorers.

<img width="2720" height="1720" alt="outlook_addin_architecture" src="https://github.com/user-attachments/assets/d5d9177c-4642-488e-a8db-8bea830160a5" />

### Models  
This add-in was tested using the free gemini-3.6-flash model – however, the program pulls data from the user's .CONFIG file, where the user enters their API key, API endpoint, and model name.

### AI API integration  
As mentioned before, the program pulls data from an app.config file that contains the AI API key, endpoint, and model name. When the user enters their instructions into a tool's form, the program sends a HTTP request with a prompt that has instructions/context based on the tool along with the user's instructions. For example, the translate email feature sends a request with the context: "You are a translation assistant for Microsoft Outlook emails. Translate the given email subject and body into {TargetLanguage}. Preserve the original meaning and tone as closely as possible. Do not add any new information, commentary, greetings, or signatures that were not already present. Do not omit any part of the original content. Respond with ONLY raw JSON, no markdown fences, no commentary, in exactly this shape: "{\\"subject\\": \\"...\\", \\"body\\": \\"...\\" ".  
The last part of the prompt allows the program to properly parse the response it receives, allowing it to correctly display the output in the form or mailItem.

<img width="2720" height="2224" alt="gemini_request_flow" src="https://github.com/user-attachments/assets/c017fa25-2ec5-4cbf-bdaa-9a368af2dcc1" />

# Tools

## Generate Email

The Generate Email form allows the user to enter the generation instructions, the tone (enforced as either casual or professional), and the maximum number of words (optional). A request with the prompt containing this information is then sent to the API along with the appropriate context included, and the program opens a new mailItem and pastes the response in the subject and body. A message box with an error message is displayed if the generation instructions or selected tone are left empty.

<img width="752" height="368" alt="Picture4" src="https://github.com/user-attachments/assets/a60f8c98-8d4f-4837-93fb-2385f4abfeb7" />

## Spell Check / Grammar Check

Upon pressing the Spell Check button, a loading form is displayed while the content of the open mailItem is sent as a request to the API along with appropriate context. After receiving the response, the contents of the mailItem are replaced with the output. If there is no mailItem open, a message box with the error message is displayed.

<img width="376" height="107" alt="Picture5" src="https://github.com/user-attachments/assets/1f9d5519-fc37-4b13-881d-a611a44100ee" />

## Reply Assist

Opens the Reply Assist form, where the user is prompted to enter the content of the email they want to reply to (this is done instead of copying the content of the open mailItem because the add-in does not support images yet - it also keeps user privacy by letting them choose what they want the tool to read.), and also the reply instructions for generation. The email content, reply instructions, and appropriate context are then sent as a request to the API, and the response is appropriately parsed to paste in the correct areas of the mailItem. A label instructs the user to keep the mail they want to reply to open so that after receiving a response, the program creates a new mailItem with the appropriate receiver email address and subject along with the generated body. If there is no mailItem open, or if either of the input boxes are empty, an error is displayed using a message box.

<img width="797" height="372" alt="Screenshot 2026-09-07 105841" src="https://github.com/user-attachments/assets/2c6f3ac5-3cf3-4dc1-81d5-53b0e35b7574" />

## Language Conversion

The Language Conversion form allows the user to strictly select from a list of 43 languages to translate TO. The program sends the content of the open mailItem as a request to the API along with the appropriate context and the target language. Upon receiving the translated text, the contents of the mailItem are replaced with the output. If there is no mailItem open, a message box with the error message is displayed.

<img width="521" height="581" alt="Picture6" src="https://github.com/user-attachments/assets/1c4318ea-cac5-452a-9508-0195f037bb45" />

## Chatbot

A simple AI chatbot that is instructed to act as an "Outlook assistant". It consists of a form that has a scrollable text window where both user and bot messages are displayed (the entire chat history). The entire chat history is saved (as a list) for the session and sent entirely to the API every time the user sends a new message in order to provide the entire context – this allows the chatbot to be as helpful as possible without "forgetting" past messages. The loading form is displayed between chat messages, and if required, the user can also clear the chat history and start a new session using the displayed button in the form.

<img width="560" height="572" alt="Picture7" src="https://github.com/user-attachments/assets/ec00c8df-85d6-4cd3-b844-97aa366b4761" />

# Error handling and data validation

Comprehensive list of errors that the program handles and displays:

- Missing API key
- Missing API endpoint
- Missing model name
- Empty/whitespace user instructions (Generate Email, Reply Assist)
- Empty/whitespace email body to check (Spell Check)
- Empty/whitespace pasted email content (Reply Assist)
- Empty/whitespace target language (Translate)
- Invalid tone selected — not "professional"/"casual" (Generate Email)
- No message to send (Chatbot)
- Request timeout: "The request to the AI service timed out."
- Can't reach the service: "Could not reach the AI service."
- HTTP 401/403: "Invalid API key"
- HTTP 429: "API quota has been exceeded"
- Any other HTTP non-success code (including 404, 503): generic "The AI API request failed ({code})."
- Response body isn't valid JSON: "unreadable response"
- No candidates/parts in the response: "no content returned"
- No active inspector / no MailItem open: "Please open an email..." (Translate, Spell Check)
- No active inspector and no Explorer selection: "Please open or select an email..." (Reply Assist)

# Extras

## Configuration

As mentioned before, an app.config file must be set up containing the API key, API endpoint, and model name. This can be made using the app.config.template file included in the repo.

## Future enhancements

Potential image processing capabilities – could be used for cases like replying to emails more appropriately or a more immersive chatbot.

## Test results

Did not face many major issues during testing apart from thread issues related to the implemented async functions – these were quickly fixed using small workarounds (non-modal dialog boxes, WindowsFormSynchronizationContext class, InvokeRequired checks, etc.)

There were also some issues with the AI requests being timed out due to heavy traffic – unfortunately the only reliable "fix" was using the model at a later time.
