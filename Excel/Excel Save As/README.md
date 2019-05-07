var fileReader = new FileReader(); 
fileReader.onload = function (e) 
{ 
alert(fileReader.result);
} 
fileReader.readAsText("C:\\Users\\Khaled\\Desktop\\table.txt"); 
