function getRequest(accessToken, apiRequest, limit, offset){    
  
  let request = UrlFetchApp.fetch(apiRequest + "?limit=" + limit + "&offset=" + offset, { 
    headers: {
      method: 'GET',
      Authorization: 'Bearer ' + accessToken
    }
  });
  
  return JSON.parse(request.getContentText());
  
}

function postRequest(accessToken, apiRequest, limit){

  let request = UrlFetchApp.fetch(apiRequest + "?limit=" + limit, { 
    headers: {
      method: 'POST',
      Authorization: 'Bearer ' + accessToken
    }
  });
  
  return JSON.parse(request.getContentText());

}