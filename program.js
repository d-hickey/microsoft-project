// The initialize function is required for all apps.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    });
}
var MyArray = [['Berlin'],['Munich'],['Duisburg']];
function writeData() {
    Office.context.document.setSelectedDataAsync(MyArray, { coercionType: 'text' });
}

function test() {
	document.getElementById("test").innerText = "yohoho: "+document.getElementById("results").innerText;
}

function ReadData() {
    Office.context.document.getSelectedDataAsync("text", function (result) {
        if (result.status === "succeeded"){
			document.getElementById("results").outerHTML = "<code id=\"results\" class=\"prettyprint\">" + result.value + "</code>";
        }
		/*
        else{
            printData(result.error.name + ":" + err.message);
        }*/
    });
	PR.prettyPrint();
}
	  
