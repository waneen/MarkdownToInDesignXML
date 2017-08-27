{
  //.mdを開く
  var mdFile = File.openDialog(".mdファイルを選択してください","*.md");
  var mdname = decodeURIComponent(mdFile.name.split(".")[0]);
  var mdPath = decodeURIComponent(mdFile);
  var htmlPath = mdPath.split(".md")[0]+".html";
  var args = ["-s -f markdown_github+simple_tables -t html -o "+htmlPath+" "+mdPath];

  //cmdを叩く
  app.doScript("C:\\Users\\edit\\AppData\\Roaming\\Adobe\\InDesign\\Version 12.0-J\\ja_JP\\Scripts\\Scripts Panel\\exe.vbs",ScriptLanguage.visualBasic,args);

  //html読み込み
  var fileObj = new File(htmlPath);
  var flag = fileObj.open("r");
  var text ="";
  if (flag == true)
  {
    text = fileObj.read();
    fileObj.close();
  }else{
    alert("ファイルが開けませんでした");
  }

  //正規表現であれする
  var body= /<body.*>([^]*?)<\/body>/g.exec(text)[1];
  body = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"+"<Root>\r\n"+"<"+mdname+">"+body+"<\/"+mdname+">"+"<\/Root>\r\n";
  var p = body.replace(/<p>/g,"<Article>")
              .replace(/<\/p>/g,"<\/Article>");
  var br = p.replace(/<br \/>/g,"");
  var code = br.replace(/<pre>.*<code>/g,"<bms>")
               .replace(/<\/code>.*<\/pre>/g,"<\/bms>");
  var table = code.replace(/<table>([^]*?)<\/table>/g, tableConvert);
  var xml =table;
  var xmlPath =mdPath.split(".md")[0]+".xml";
  saveXml(xmlPath,xml);

  //xmlの読み込み
  var docObj = app.activeDocument;
  xmlPath = xmlPath.substr(1);
  driveLatter = xmlPath.split("/")[0]+"://";
  xmlPathes = xmlPath.split("/");
  xmlPathes.shift();
  xmlPath = driveLatter+xmlPathes.join("/");
  docObj.importXML(xmlPath);
}

function tableConvert(matchPart,text){
  var cells=new Array();
  var rowNum=0;
  var colNum=0;
  //head
  var head = text.match(/<thead>([^]*?)<\/thead>/g)[0];
  var headCells = head.match(/<th>([^]*?)<\/th>/g);
  for(var i = 0;i<headCells.length;i++){
    cells.push("<Cell aid:table=\"cell\" aid:theader=\"\" aid:crows=\"1\" aid:ccols=\"1\">"+
              "<strong>"+/<th>([^]*?)<\/th>/g.exec(headCells[i])[1]+"<\/strong>"+
              "<\/Cell>\r\n");
  }
  rowNum= headCells.length;
  //body
  var body = text.match(/<tbody>([^]*?)<\/tbody>/g)[0];
  var bodyCells = body.match(/<td>([^]*?)<\/td>/g);
  for(var i = 0;i<bodyCells.length;i++){
    cells.push("<Cell aid:table=\"cell\" aid:crows=\"1\" aid:ccols=\"1\">"+
              /<td>([^]*?)<\/td>/g.exec(bodyCells[i])[1]+
              "<\/Cell>\r\n");
    colNum++;
  }
  colNum=colNum/rowNum+1;
  var result ="<story xmlns:aid5=\"http://ns.adobe.com/AdobeInDesign/5.0/\" xmlns:aid=\"http://ns.adobe.com/AdobeInDesign/4.0/\">\r\n"+
              "<Table aid:table=\"table\" aid:trows=\""+colNum+"\" aid:tcols=\""+rowNum+"\">\r\n"+
              cells.join("")+
              "<\/Table>\r\n"+
              "<\/story>\r\n";

  return result;
}

function saveXml(filename,text){
  fileObj=new File(filename);
  fileObj.encoding = "UTF-8";
  flag = fileObj.open("w");
  if (flag == true)
  {
    fileObj.write(text);
    fileObj.close();
  }else{
    alert("ファイルが開けませんでした");
  }
}
