myObject = new ActiveXObject("Scripting.FileSystemObject");
f = myObject.GetFolder(img);
console.log(f.files.Count)
