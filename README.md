Название сервиса: support-automation <br>
API получает от Google Apps script токен, обновленное название столбца, новое значение столбца, postID и информацию о том, является ли это чат менеджера или преподавателя. <br>
Затем служба обмена сообщениями отправит соответствующее сообщение (MonitorService.cs -> SendUpdateMessage).<br>

[MonitorService.cs] Сервис, который активно отслеживает новые добавленные строки и обрабатывает логику отправки по соответствующим каналам.

<br> <br> Web service that gets real-time updates from Google Sheet (new rows added and edits) and synchronize with Mattermost. 

<br>

<h2> How it works: </h2> <br>
- Web service parse for newly added rows <br>
- If a row is added in one of the sheets (Managers/Teachers) it sends a post to Mattermost channel with the according information <br>
- It then adds the PostId to the row in the sheet <br>
- When an update occurs on a row, the Google Apps Script will send the update with the PostId to the API service so that it transfers the information to Mattermost channel <br>

<h4>Note: The service is not fully optimized, you can add secrets on the appsettings file and access it using IConfiguration for better security.  <br> Also, the parsing algorithm can be improved. <br> Logging using ILogging or other logging libraries for production </h4>
