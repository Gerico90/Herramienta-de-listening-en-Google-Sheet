function sheet(n){   
    var s =  SpreadsheetApp.openById("ID del docuemnto aquí");
     var ss = s.getSheetByName('Registro');
let ultimaFilaDestino = ss.getRange("A:A").getLastRow();
       var sss = ss.getRange(ultimaFilaDestino+1,1,1,4).setValues(n);  
return sss;
}

function user(){
     var correo = Session.getActiveUser().getEmail(); // Correo
        var name = correo.replace('@grupow.com',''); // Extracción Nombre
          var user = name.charAt(0).toUpperCase() + name.slice(1); // Usuario
return user;
}

function buscador(){

var app = SpreadsheetApp;
var ss = app.getActiveSpreadsheet();
var activesheet = ss.getSheetByName("Buscador");

function formatDate(date) {
   var d = new Date(date),
       month = '' + (d.getMonth() + 1),
       day = '' + d.getDate(),
       year = d.getFullYear();

   if (month.length < 2) 
       month = '0' + month; 
   if (day.length < 2) 
       day = '0' + day;
   return [year, month, day].join('-');
}

var fechaInicio = formatDate(activesheet.getRange("C5").getValue())
var fechaFinal = formatDate(activesheet.getRange("C6").getValue())
var maximo = activesheet.getRange(7,2).getValue()
var tipo = activesheet.getRange(9,2).getValue()
var rtweet = activesheet.getRange("F6").getValue()
var reply = activesheet.getRange("G6").getValue()
var query1 =activesheet.getRange("C2").getValue()
var query2 = " " + activesheet.getRange("C3").getValue()

if(rtweet==true){
 var rtweet =" -is:retweet"
}else{
 var rtweet =""
}

//Logger.log (rtweet)

if(reply==true){
 var reply =" -is:reply"
}else{
 var reply =""
}
//Logger.log (reply)

var query= query1+query2
var query = encodeURIComponent(query)
//var query = encodeSolicitud
//Logger.log(solicitud)
//Logger.log (query);

var url ="https://api.twitter.com/2/tweets/search/recent?query="+query+rtweet+reply+"&max_results=100"+"&end_time="+fechaFinal+ "T00:00:00-15:00" +"&start_time="+fechaInicio+"T00:00:00.00Z"+"&tweet.fields=created_at,public_metrics&expansions=author_id,attachments.media_keys&media.fields=url,preview_image_url";

//Logger.log(url)

const res = UrlFetchApp.fetch(url,{
      muteHttpExceptions: true,
         headers:{
           Authorization: "Bearer AGREGAR EL TOKEN DE ACCESO"
         },
       })

var responseCode = res.getResponseCode();

Logger.log (responseCode)
Logger.log (res)


var datos={};

if(responseCode !== 200){
 SpreadsheetApp.getUi().alert("Este es un error " + responseCode + ", reportelo con Jesús García");
}else{

 datos = JSON.parse(res.getContentText())}
 //ui.alert('Hello, world');
 //Logger.log(datos.meta.result_count)
 let huboResultados= true 
 if(datos.meta.result_count == 0){ 
   huboResultados = false
   SpreadsheetApp.getUi().alert("No hubo resultados, verifique su query e intente nuevamente")
 }
let mainArray =[];

if(huboResultados){
const values = () => {
   const {data, includes} = datos
   let arrayData = [] 
   let arrayMetricas = []

   data.map(item => {

       

       const { author_id, id, text, created_at, public_metrics} = item 
       let {attachments} = item
       //const {media_keys} = attachments
       if(attachments == null){
         attachments = {media_keys:["000000"]}
       }

       const attachments_key = attachments.media_keys[0];

       Logger.log(attachments_key)

       const {retweet_count, reply_count, like_count} = public_metrics;
       const totalinteracciones = retweet_count + reply_count + like_count 
       const idTweet = id;

       Logger.log(public_metrics)
       let urlmio = "";

       includes.media.map(media => {
         const {media_key,url,preview_image_url,type} = media

         if(media_key == attachments_key){
             if(type == 'video'){
                 urlmio = "=IMAGE("+'"'+preview_image_url+'"'+",2)"
             }if(type == 'photo'){                
                urlmio = "=IMAGE("+'"'+url+'"'+",2)"
             }
         }/*else{

           urlmio = "vacio"

         }*/

       })

       includes.users.map(user => {
           const { id, name, username} = user

           if(author_id === id){
               arrayData = ["twitter.com/anyuser/status/" + idTweet, formatDate(created_at), username, text, urlmio, like_count, reply_count, retweet_count, totalinteracciones]
           }
           return arrayData
       })
       Logger.log(arrayData)

       return mainArray.push(arrayData)
   })   
}

//Logger.log(mainArray)

values()

var totaltweets = mainArray.length


var fechaHoy = new Date();
var correo = Session.getActiveUser().getEmail();
//let registro = [[fechaHoy], [url], [correo], [totaltweets]]
let registro = [[fechaHoy, url, correo, totaltweets]]
sheet(registro)

//Logger.log(query)
//Logger.log(totaltweets)
//Logger.log (activesheet.getLastRow())
activesheet.getRange("B10:L"+ activesheet.getLastRow()).clearContent()
//Logger.log(mainArray[0].length)
activesheet.getRange(10,2,totaltweets, mainArray[0].length).setValues(mainArray);
}

//let urlImages = activesheet.getRange("J10:J").getValue();

}