# Get your course codes

Generate your course registration slip from [CollPoll](https://bennett.digiicampus.com/courseRegistration/student).

Slip looks like this
![Course Registration](image.png)

Open Dev Tools and run the following code in the console window.
```JS
// Amatuer JS Codes.
table = document.childNodes[1].childNodes[1].childNodes[1].childNodes[0].childNodes[0].childNodes[4].childNodes[2].childNodes[9].childNodes[1].childNodes[1]

f=''
for(var i=2; i < $0.childNodes.length-2; i+=2) {
   f += table.childNodes[i].cells[1].innerText + '\n'
}
console.log(f)
```