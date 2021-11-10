# EmailToExcelConsole
<div>
<ul>
	<li>C#</li>
	<li>.NET Framework 4.7.2</li>
	<li><a href="https://epplussoftware.com/docs/5.0/api/OfficeOpenXml.html">OfficeOpenXml</a></li>
	<li><a href="https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data?view=exchange-ews-api">Microsoft.Exchange.WebServices.Data</a></li>
	
</ul>
</div>


Console program to read specific messages from office365. With default settings this console program will read 100 mails from office365 account and write only those mails where topic's first 10 charachters contains date. The excel file will contain the following info:
<ul>
<li>Date Recieved</li>
<li>Date TODO</li>
<li>Message Sender</li>
<li>Message Contents (plain text only)</li>
</ul>

## Usage
<p>User must enter Office365 credentials within the code.</p>
<p>User is able to specify <b>mailsToRead</b> in the code (default = 100).</p>

## Screenshots
<a data-flickr-embed="true" href="https://www.flickr.com/photos/55156353@N07/51669040602/in/dateposted-public/" title="excelFrontEndConsole"><img src="https://live.staticflickr.com/65535/51669040602_d0301e73f2_b.jpg" width="979" height="512" alt="excelFrontEndConsole"></a>
<br />
<a data-flickr-embed="true" href="https://www.flickr.com/photos/55156353@N07/51669838186/in/dateposted-public/" title="excelTest"><img src="https://live.staticflickr.com/65535/51669838186_9f3e5c7fbd_z.jpg" width="640" height="113" alt="excelTest"></a>
