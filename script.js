function doGet() {
    return HtmlService.createHtmlOutputFromFile('sidebar');
  }
  
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu("Custom GPT")
      .addItem("GPT Function", "gptSidebar")
      .addToUi();
  }
  
  function gptSidebar(){
    var html = HtmlService.createHtmlOutputFromFile("sidebar").setTitle("Test1");
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  function receiveDataValueModel(tempValueModel){
    PropertiesService.getUserProperties().setProperty('CHATGPT_VERSION', tempValueModel);
  }
  
  function receiveDataToken(token){
    PropertiesService.getUserProperties().setProperty('TOKEN_INTEGER_VALUE', token.toString());
  }
  
  function receiveDataTemperature(temperature){
    var temperatureFloat = parseFloat(temperature).toFixed(2);
    PropertiesService.getUserProperties().setProperty('SYSTEM_TEMPERATURE', temperatureFloat);
  }
  
  function receiveDataSavedPrompt(systemMessage){
    PropertiesService.getUserProperties().setProperty('SYSTEM_MESSAGE', systemMessage);
  }
  
  function receiveDataPreprompt(prePrompt){
    PropertiesService.getUserProperties().setProperty('SYSTEM_PREPROMT', prePrompt);
  }
  
  function deleteDataSavedRole(){
    PropertiesService.getUserProperties().setProperty('SYSTEM_MESSAGE', "You are the best assistance ");
  }
  
  function deleteDataPreprompt(){
    PropertiesService.getUserProperties().setProperty('SYSTEM_PREPROMT', " ");
  }
  
  function getDynamicKey(keyFromInput){
    PropertiesService.getUserProperties().setProperty('SYSTEM_API_KEY', keyFromInput);
  }
  
  function callGPT(cell) {
    var concatenatedTextPrompt = '';
  
    for (var row = 0; row < cell.length; row++) {
      for (var col = 0; col < cell[row].length; col++) {
        concatenatedTextPrompt += cell[row][col] + " ";
      }
    }
    var requestData = {
      //model: "gpt-3.5-turbo",
      //model: "gpt-4",
      //model: "gpt-4-turbo-preview",
      model: PropertiesService.getUserProperties().getProperty('CHATGPT_VERSION'),
      max_tokens: parseInt(PropertiesService.getUserProperties().getProperty('TOKEN_INTEGER_VALUE')),
      temperature: parseFloat(PropertiesService.getUserProperties().getProperty('SYSTEM_TEMPERATURE')),
      messages: [
        {
          role: "system",
          content: PropertiesService.getUserProperties().getProperty('SYSTEM_MESSAGE'),
        },
        {
          role: "user",
          content: PropertiesService.getUserProperties().getProperty('SYSTEM_PREPROMT') + concatenatedTextPrompt
        }
      ]
    };
  
    // Set up the API endpoint and your API key
    var apiKeyFromUserProperty = PropertiesService.getUserProperties().getProperty('SYSTEM_API_KEY');
    var url = 'https://api.openai.com/v1/chat/completions';
    var apiKey = apiKeyFromUserProperty; //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
  
    // Set the options for the request  
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      // Convert the JavaScript object to a JSON string
      'payload': JSON.stringify(requestData),
      'headers': {
        'Authorization': 'Bearer ' + apiKey
      },
      'muteHttpExceptions': true // To handle HTTP errors more gracefully
    };
  
    // Make the request and handle the response
    try {
      var response = UrlFetchApp.fetch(url, options);
  
      // Check for response code to see if the request succeeded
      if (response.getResponseCode() !== 200) {
        throw new Error("HTTP error! status: " + response.getResponseCode());
      }
      Logger.log(response);
  
      // Parse the JSON response
      var data = JSON.parse(response.getContentText());
  
      // Log the entire response
      console.log('Response from OpenAI:', data);
      var result = data.choices[0].message.content;
      Logger.log(result);
  
    } catch (error) {
      // Log any errors
      console.error('There was an error!', error.toString());
    }
    return result;
  }