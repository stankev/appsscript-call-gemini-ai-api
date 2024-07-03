/******************************************************************************************************
 *
 * Name:                 callGeminiAPI
 * Description:          This Google Apps script is intended to be used to demonstrate how to call  
 *                       the Google Vertex AI Gemini API from a Google doc using Apps scripts.
 *
 *                       The code will extract the meeting minutes from the Google doc, generate a 
 *                       formatted list of action items extracted from the meeting minutes, and 
 *                       insert them into the google doc.
 * 
 *                       The base code can be used to make other calls to Gemini API with changes to  
 *                       the prompt and the parameters passed to Gemini.                     
 *                       
 * Date:                 June 17, 2024
 * Author:               Mark Stankevicius
 * GitHub Repository:    https://github.com/stankev/appsscript-call-gemini-ai-api
 *
 *******************************************************************************************************
 */
 function callGeminiAPI() {

  // Define the constants for the Gemini Vertex API call
    
    const projectId = "YOUR_PROJECT_ID";       // enter your project Id from the google cloud, or for more security use a script properties to store the ID
    const modelId = 'gemini-1.5-pro-001';               // choose one of the Gemini models you want to use for this script
    const location = 'us-central1';                     // choose the location for the API resources 
    const apiURL = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${modelId}:generateContent`; 
  
    // define the directions to the API on what action to take and how to generate the output
    const SYSTEM_PROMPT = `Simulate three brilliant, logical project managers reviewing status meeting minutes and determining the action items.
      The three project managers review the provided status meeting minutes and create their list of action items. The three project managers must 
      carefully review the full list of meeting notes to ensure they capture any action items hidden in the minutes. 
      The three experts then compare their list against the action item list from the other project managers. 
      Based on their comparison they generate a final list of action items. Do not generate action items for milestones in the minutes.
      The action items should list the following for each item: title of the action item, due date if known, dependency if known, and owner of the action item. 
      If there is no dependency, then state None.  Try to infer the due date based on the minutes, but if the due date cannot be determined than specify TBD. 
      Be accurate when creating the action items and do not make up fictitious action items that are not in the meeting minutes.
      Format the output so that it could be inserted into a google doc. Do not bold text or any special formatting such as asterisks for emphasis.
      An example format of output is the following:
      Action Item: Develop Power User training plan and materials.\n - Due Date: April 25/2024\n  - Dependency: Completion of Teams training materials.\n - Owner: Alice Williams (Training Lead)`;
  
    const extText = extractText();    // extract the text from the document
   
    // Validate text was found in the document. If no text is send a message to the log
    if (!extText.success) {         // Test whether text was extracted 
      Logger.log('A problem occurred with the input document');
      Logger.log(`No text was found. Error code: ${extText.error.code}\nError message: ${extText.error.message}`);
      return;
    }
  
    // If successful, get the text from the result object
    const chatPrompt = extText.extractResults; 
  
    // Construct Request Payload with the meeting minutes as input 
    const payload = {
      contents:[
        {
        role: 'USER',
        parts: [{ text: chatPrompt }]
        }
      ],
      system_instruction: {
        parts: [{ text: SYSTEM_PROMPT }]
      },
      generation_config: {
        temperature: 0.2, 
        topP: 0.3,
        candidateCount: 1,
        maxOutputTokens: 800
      }
    };
  
    // Set Request Options - use the OAuth token for authorization to the API
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  
    // call the API to review the meeting minutes and generate action items
    try {
      const response = UrlFetchApp.fetch(apiURL, options);               // Call to the Vertex AI API with the Gemini model
  
      if (response.getResponseCode() === 200) {
        
        // parse the response from Gemini AI API into a JSON object
        const chatResponseJson = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;  
        
        // Validate the response was successful and if successful append action items to document
        if (chatResponseJson && chatResponseJson.trim().length > 0) {           // is there generated text in the response
              
          // write a heading line in the document
          const document = DocumentApp.getActiveDocument();
          const body = document.getBody();
          body.appendParagraph('Meeting minutes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        
          // Append the returned action items to the current active document
          body.appendParagraph(chatResponseJson);
  
        } else {
          Logger.log('No action items were generated');
          return;
        } 
      } else {
          console.error('An error occurred on the Vertex AI API call');   // log an error condition in the execution log
          const errorCode = response.getResponseCode();
          if (errorCode === 404) {
            console.error('Code: ' + errorCode + ' - API URL is undefined');       // log the error code            
            console.error('URL: ' + apiURL);
          } else {
            console.error('Code: ' + errorCode);                            // log the error code
            const errorResponse = JSON.parse(response.getContentText());    // parse the response object
            console.error('message: ' + errorResponse.error.message);       // log the error message
            console.error(response.getContentText());
          }
      }

    } catch (error) {                                             // An unexpected error occurred in the code, therefore log the details as an error
        console.error('An unexpected error occurred:');           // log any errors to the apps script logger
        console.error('Error name: ' + error.name);               // log the error name 
        console.error('Error message: ' + error.message);         // log the error message 
        if (error.stack) {
          console.error('Stack Trace: ' + error.stack);           // if the error stack information is available log it to the execution log
        }
    }
  }
  
    
  /*
   * Create a custom menu in the Google Docs UI for the "AI Tools" menu with one item: "Generate Action Items".
   * When the "Generate Action Items" item is clicked, it calls the "callGeminiAI" function.
   */
  function onOpen() {
    const ui = DocumentApp.getUi();               // Get a reference to the user interface object for the Google Docs document.
    // Create a new menu in the Google Docs UI with the title "Gemini AI Tools".
    const menu = ui.createMenu('AI Tools')
                   .addItem('Generate Action Items', 'callGeminiAPI');
    menu.addToUi();
  }
  
  /*  
   * Function to extract text from the document and return it to the caller
   */
  function extractText() {
    
    const document = DocumentApp.getActiveDocument();
    const body = document.getBody();
    const text = body.getText();
  
    // If there is text found in the document return it to the caller
    if (text.trim().length > 0) {                   // check if there is text in the document
        return {
          success: true,                      // return success = true, when the meeting minutes were found
          extractResults: text.trim()         // assign the meeting minutes to the return results
        };
    } else {                                  // The start and end delimiters were not found - send an error code indicating missing delimiters
        return {
        success: false,                       // return success = false, when the document is empty
        error: {                              // return the error information to indicate which error occurred 
          message: 'No text was found in the document.',
          code: 101
        }
      }
    }
  }
