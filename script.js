const theadRow = document.getElementById("table-heading-row");
const tbody = document.getElementById("table-body");
const boldButton = document.getElementById("bold-btn");
const italicButton = document.getElementById("Italic-btn");
const underlineButton = document.getElementById("underline-btn");
////
const textColor = document.getElementById("text-color");
const bgColor = document.getElementById("bg-color");
////
const leftAlign = document.getElementById("left-align");
const centerAlign = document.getElementById("center-align");
const rightAlign = document.getElementById("right-align");
////
const fontSizeValue = document.getElementById("font-size");
const fontFamilyValue = document.getElementById("font-family");
////
const cutButton = document.getElementById("cut");
const copyButton = document.getElementById("copy");
const pasteButton = document.getElementById("paste");



let cutValue = {};
let currentCell;

//rows
const rows = 100;
//cols
const cols = 26;



let matrix = new Array(rows);

for (let i = 0; i < rows; i++) {
    matrix[i] = new Array(cols);
    for (let j = 0; j < cols; j++) {
        matrix[i][j] = {};
    }
}

function updateMatrix(currentCell) {
    let obj = {
        style: currentCell.style.cssText,
        text: currentCell.innerText,
        id: currentCell.id
    }

    let id = currentCell.id.split("");
    //If we go horizontally
    // id = ['A' ,1] --> 0,0 
    // id = ['B' ,1] --> 0,1
    // id = ['C' ,1] --> 0,2

    //If we go vertically
    // id = ['A' ,1] --> 0,0 
    // id = ['A' ,2] --> 1,0
    // id = ['A' ,3] --> 2,0

    let i = id[1] - 1;
    let j = id[0].charCodeAt(0) - 65;

    matrix[i][j] = obj;
}


//for adding heading row in sheet
for (let col = 65; col <= 90; col++) {
    let th = document.createElement("th");
    th.innerText = String.fromCharCode(col);
    theadRow.append(th);
}

//for 
for (let row = 1; row <= 100; row++) {
    let tr = document.createElement("tr");
    let th = document.createElement("th");
    th.innerText = row;
    tr.appendChild(th);
    //Looping from A to Z
    for (let col = 1; col <= 26; col++) {
        let td = document.createElement("td");
        td.setAttribute("contenteditable", "true");

        //colRow --> A1 , B1 , C!...
        td.setAttribute("id", `${String.fromCharCode(col + 64)}${row}`);
        //this is a event Listner
        td.addEventListener("focus", (event) => onFocusFunction(event));
        td.addEventListener("input", (event) => onInputFunction(event));
        tr.appendChild(td);
    }
    tbody.appendChild(tr);
}

//defining things related to my sheets
let numSheets = 1 ;
let arrMatrix = [matrix] ;
let currSheetNum = 1 ; 

//onfocus function
function onFocusFunction(event) {
    // console.log(event.target);
    currentCell = event.target;
    document.getElementById("current-cell").innerText = currentCell.id;
}

// on input function
function onInputFunction(event) {
    updateMatrix(event.target);
}

// Download Function
function downloadJson(){
    // console.log("kuch bhi") ;
    const jsonString = JSON.stringify(matrix) ;

    const blob = new Blob([jsonString],{type : "application/json"}) ;

    const link = document.createElement('a') ;
    link.href = URL.createObjectURL(blob) ;

    // this is file name
    link.download = "data.json" ;

    document.body.appendChild(link) ;
    link.click() ;
    document.body.removeChild(link) ;
    
}

// Upload Functionality /// readJsonFile is a function
document.getElementById("jsonFile").addEventListener("change" , readJsonFile) ;

function readJsonFile(event){

    // here we got our file
    const file = event.target.files[0] ;

    if(file){
        // reader object which is instance of FileReader 
        const reader = new FileReader() ;

        reader.onload = function(e){
            const fileContent = e.target.result ;
            try{
                const jsonData = JSON.parse(fileContent) ;

                //we are not doing anything related to upload
                matrix = jsonData ;
                jsonData.forEach((row)=>{
                    // cell is cell inside matrix (virtual excel)
                    row.forEach((cell)=>{
                        if(cell.id){
                            // myCell is cell inside my DOM or reel Excel
                            var myCell = document.getElementById(cell.id) ;
                            myCell.innerText = cell.text ;
                            myCell.style.cssText = cell.style ;
                        }
                    }) ;
                }) ;

            }catch (err) {
                console.log("Error in reading Json file"+ err) ;
            }
        }
        reader.readAsText(file) ;
    }
}


//BOLD Button
boldButton.addEventListener("click", () => {

    if (currentCell.style.fontWeight == "bold") {
        currentCell.style.fontWeight = "normal";
    }
    else {
        currentCell.style.fontWeight = "bold";
    }

    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//ITALIC Button
italicButton.addEventListener("click", () => {
    if (currentCell.style.fontStyle == "italic") {
        currentCell.style.fontStyle = "normal";
    } else {
        currentCell.style.fontStyle = "italic";
    }
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//UNDERLINE Button
underlineButton.addEventListener("click", () => {
    if (currentCell.style.textDecoration == "underline") {
        currentCell.style.textDecoration = "none";
    } else {
        currentCell.style.textDecoration = "underline";
    }
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//Text Color
textColor.addEventListener("click", () => {
    currentCell.style.color = textColor.value;
    // I need to pass the updated current cell Inside updatedMatrix 
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//Bg-Color
bgColor.addEventListener("click", () => {
    currentCell.style.bgColor = bgColor.value;
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//Left Align Button
leftAlign.addEventListener("click", () => {
    currentCell.style.textAlign = "left";
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//Right Align Button
rightAlign.addEventListener("click", () => {
    currentCell.style.textAlign = "right";
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//Center Align Button
centerAlign.addEventListener("click", () => {

    currentCell.style.textAlign = "center";

    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

// Font -Size
fontSizeValue.addEventListener("change", () => {
    currentCell.style.fontSize = fontSizeValue.value;

    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
})

//Font-Family
fontFamilyValue.addEventListener("change", () => {
    currentCell.style.fontFamily = fontFamilyValue.value;
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
})

//Cut Button
cutButton.addEventListener("click", () => {
    cutValue = {
        style: currentCell.style.cssText,
        text: currentCell.innerText,
    };
    currentCell.style = null;
    currentCell.innerText = "";
    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
});

//Copy Button
copyButton.addEventListener("click", () => {
    cutValue = {
        style: currentCell.style.cssText,
        text: currentCell.innerText,
    };

});

//Paste Button
pasteButton.addEventListener("click", () => {
    if (cutValue.text) {
        currentCell.style = cutValue.style;
        currentCell.innerText = cutValue.text;

    // I need to pass the updated current cell Inside updatedMatrix
    updateMatrix(currentCell) ;
    }
});


// add sheet button 
document.getElementById("add-sheet-btn").addEventListener("click", () => {
    // logic for adding sheet
  
    if(numSheets <3){

    
    /// logic for saving prevSheets
    if (numSheets == 1) {
      var myArr = [matrix];
      localStorage.setItem("ArrMatrix", JSON.stringify(myArr));
    } else {
      var prevSheets = JSON.parse(localStorage.getItem("ArrMatrix"));
      var updatedSheets = [...prevSheets, matrix];
      localStorage.setItem("ArrMatrix", JSON.stringify(updatedSheets));
    }
}else{
    alert("You can only add 3 sheets"); 
}
    ///updateMy number of sheets
    numSheets++;
    currSheetNum = numSheets;
  
    // cleanup my virtual memory of excel which is matrix;
    for (let i = 0; i < rows; i++) {
      matrix[i] = new Array(cols);
      for (let j = 0; j < cols; j++) {
        matrix[i][j] = {};
      }
    }
  
    // cleaning up excel sheet in HTML;
    tbody.innerHTML = ``;
  
    for (let row = 1; row <= 100; row++) {
      let tr = document.createElement("tr");
      let th = document.createElement("th");
      th.innerText = row;
      tr.appendChild(th);
      // looping from A to Z;
      for (let col = 1; col <= 26; col++) {
        let td = document.createElement("td");
        td.setAttribute("contenteditable", "true");
        // colRow -> A1, B1, C1, D1,
        td.setAttribute("id", `${String.fromCharCode(col + 64)}${row}`);
        // this is the event listener
        td.addEventListener("focus", (event) => onFocusFunction(event));
        td.addEventListener("input", (event) => onInputFunction(event));
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
  
    document.getElementById("sheet-num").innerText = "Sheet No." + currSheetNum ;
})

document.getElementById("sheet-1").addEventListener("click",  () => {
    document.getElementById("sheet-num").innerText = "Sheet No." + 1 ;
    var myArr = JSON.parse(localStorage.getItem("ArrMatrix"));
    let tableData = myArr[0];
    matrix = tableData;
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (cell.id) {
          var myCell = document.getElementById(cell.id);
          myCell.innerText = cell.text;
          myCell.style.cssText = cell.style;
        }
      });
    });
  });

document.getElementById("sheet-2").addEventListener("click", () => {
    document.getElementById("sheet-num").innerText = "Sheet No." + 2 ;
    var myArr = JSON.parse(localStorage.getItem("ArrMatrix"));
    let tableData = myArr[1];
    matrix = tableData;
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (cell.id) {
          var myCell = document.getElementById(cell.id);
          myCell.innerText = cell.text;
          myCell.style.cssText = cell.style;
        }
      });
    });
  });

document.getElementById("sheet-3").addEventListener("click", () => {
    document.getElementById("sheet-num").innerText = "Sheet No." + 3 ;
    var myArr = JSON.parse(localStorage.getItem("ArrMatrix"));
    let tableData = myArr[2];
    matrix = tableData;
    tableData.forEach((row) => {
      row.forEach((cell) => {
        if (cell.id) {
          var myCell = document.getElementById(cell.id);
          myCell.innerText = cell.text;
          myCell.style.cssText = cell.style;
        }
      });
    });
  });

