// The initialize function is required for all apps.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    });
}


function parseCode(code) {
	//document.getElementById("debug").outerHTML = "gets into function";
    var result = "";
    var pos = 0;
    var c;
    var indent = 0;
	var newline = 0;
	var infor = 0;
    code = removeExtraSpaces(code);
	
    while(pos < code.length) {
        c = code.charAt(pos);
        if (infor === 0){
			if (c === '}'){
				indent = indent - 1;
			}
			if(newline === 1){
				result = result + makeIndent(indent);
				newline = 0;
			}
			
			result = result + c;
		
        
			if (c === '{'){
				indent = indent + 1;
				while(code.charAt(pos+1) === ' '){
					pos = pos + 1;
				}
				result = result + '\n';
				newline = 1;
			}
			if (c === ';'){
				while(code.charAt(pos+1) === ' '){
					pos = pos + 1;
				}
				result = result + '\n';
				newline = 1;
			}
			// this part doesn't seem to work too well atm
			/*if (code.substring(pos, pos+3) == "for("){
				infor = 1;
			}*/
			
			
		}
		else{
			result = result + c;
			if(code.charAt(pos+1) === '{'){
				infor = 0;
			}
		}
  
        pos = pos + 1;
    }
	//document.getElementById("debug").outerHTML = "finishes parse function";
    return result;
}

function removeExtraSpaces(code){
    var result = "";
    var seenSpace = 0;
    var c;
    for(var i = 0; i < code.length; i++){
        c = code.charAt(i);
        if(c !== '\n'){
            if(c !== ' ' || seenSpace != 1){
				seenSpace = 0;
                result = result + c;
                if(c === ' '){
                    seenSpace = 1;
                }
            }
        }
    }
    return result;
}

function makeIndent(ind) {
    var space = "";
    var i = 0;
    while(i < ind){
        space = space + "    ";
        i++;
    }
    return space;
}

function test() {
	document.getElementById("test").innerText = "yohoho: "+document.getElementById("results").innerText;
}

function ReadData() {
	var code = document.getElementById('code').value;
	
   /* parseCode(code, function (result) {
        if (result.status === "succeeded"){
			document.getElementById("results").outerHTML = "<code id=\"results\" class=\"prettyprint\">" + result.value + "</code>";
			Office.context.document.setSelectedDataAsync(result.value, { coercionType: 'text' });
        }
        else{
            printData(result.error.name + ":" + err.message);
        }
    });*/
	document.getElementById("debug").outerHTML = "read code: "  + code;
	var formatted = parseCode(code);
	document.getElementById("deb2").outerHTML = "Passed Formatting";
	document.getElementById("results").outerHTML = "<pre id=\"results\" class=\"prettyprint\">" + formatted + "</pre>";
	Office.context.document.setSelectedDataAsync(formatted, { coercionType: 'text' });
	PR.prettyPrint();
}
	  
