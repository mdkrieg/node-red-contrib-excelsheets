module.exports = function (RED) {
    "use strict";
    var json2xl = require('json2excel');

    function isObject(input_obj){
        return (typeof input_obj === "object" &&  input_obj != null);
    }
    
    function ConvertJSONToExcel(config) {
        RED.nodes.createNode(this, config);
        this.on('input', function(msg){
            var node = this;
            var file = msg.filepath || config.file;
            if(file === undefined || file === null ){
                node.error("filepath cannot be null");
                return;
            }
            if(msg.payload === undefined || msg.payload === null){
                node.error("payload cannot be null");
                return;
            }
            if(isObject(msg.payload)){
                msg.payload = convertObjectToArray(msg.payload);
            }
            json2xl.j2e({
                sheets      :msg.payload,
                filepath    :file
            },function(){
                node.send(msg);
            });
        });
    }
    function convertObjectToArray(obj){
        let output = [];
        
        // LEVEL 3+ - disseminate arbitrary objects
        function handleRowObject(row_obj, prefix = []){
            let output_row = {};
            // handles rows as well as "layered" objects inside the rows
            //   the prefix variable is used to carry through the path
            for(var key in row_obj){
                let data = handleDataItem(row_obj, key, prefix);
                // merge the handled data to our output row
                if(Array.isArray(data)){
                    // returns array if primitive item
                    // NOTE: every branch eventually gets here
                    output_row[data[0]] = data[1];
                }else if(isObject(data)){
                    // returns object if complex item
                    output_row = {...output_row, ...data};
                }
            }
            return output_row;
        }
        function handleDataItem(parent_obj, data_key, prefix){
            // looks if it's an object then it recurses through
            let content = parent_obj[data_key];
            let full_prefix = [...prefix, ...[data_key]];
            if(isObject(content)){
                return handleRowObject(content, full_prefix);
            }
            return [full_prefix.join("-"), content];
        }
        
        // LEVEL 1 - object keys are sheet names
        for(var sheet in obj){
            let sheet_obj = {
                header: {},
                items: [],
                sheetName: sheet
            };
            // LEVEL 2 - keys contain arrays of data
            if(!Array.isArray(obj[sheet])) continue;
            for(var row of obj[sheet]){
                var flat_row = handleRowObject(row);
                for(var key_name in flat_row){
                    // headers are just simply the key names
                    // this method is conveniently passive on duplicates
                    sheet_obj.header[key_name] = key_name;
                }
                sheet_obj.items.push(flat_row);
            }
            output.push(sheet_obj);
        }
        return output;
    }
    RED.nodes.registerType("excelsheets", ConvertJSONToExcel);
}
