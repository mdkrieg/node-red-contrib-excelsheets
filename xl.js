module.exports = function (RED) {
    "use strict";
    var json2xl = require('json2excel');

    function ConvertJSONToExcel(config) {
        
        RED.nodes.createNode(this, config);
        this.file = config.file;
        this.on('input', function(msg){
            var node = this;
            var file = msg.filepath;
            if(file === undefined || file === null){
               if(node.file === undefined || node.file === null ){
                   console.log("File path cannot be null");
                   node.error("filepath cannot be null");
               }else{
                   file = node.file;
               }
            }
            if(msg.payload === undefined || msg.payload === null){
                console.log("Payload cannot be blank");
                node.error("payload cannot be null");
            }else if(!Array.isArray(msg.payload)){
                console.log("Payload must be Array");
                node.error("payload must be Array");
            }else{
                json2xl.j2e({
                    sheets      :msg.payload,
                    filepath    :file
                },function(){
                    node.send(msg);
                });
            }
        });
        
    };

    RED.nodes.registerType("excelsheets", ConvertJSONToExcel);
}
