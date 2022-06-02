function getContractors(guidArr) {
  
  let service = getDriveService();
  let accessToken = service.getAccessToken()

  let request = UrlFetchApp.fetch("", {
    method:"post", 
    headers:{
      "accept": "text/plain",
      "Authorization" : "Bearer " + accessToken,
      "Content-Type" : "application/json-patch+json"
    },
    payload: JSON.stringify({
      "guid1C": guidArr
    })
  });  

  return JSON.parse(request.getContentText()).data.map(obj => [obj.name, obj.inn, obj.kpp, obj.guid1C]);

}
