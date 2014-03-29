// The initialize function is required for all apps.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    });
}

//gogo global variable
var level=0;
var LOOP_SIZE=100;
var tabstop = 4;
var type = 'Java';
var highlight = 0;
var format = 0;

function changeHide(one, two){
    document.getElementById(one).style.display = 'none';
    document.getElementById(two).style.display = 'block';
}

function changeLang(lang){
    type = lang;
}

function changeIndent(newindent){
    tabstop = newindent;
}

function runTabifier() {
  //alert("tabifier runs");
  var code = document.getElementById('code').value; 
  
  //alert(document.getElementById("colour").checked);
  highlight = 0;
  if(document.getElementById("colour").checked === true){
    highlight = 1;
  }
  format = 0;
  if(document.getElementById("format").checked === true){
    format = 1;
  }
  
  //console.log(tabstop+"ok");
  if(format === 1){
      if ('C'==type) code=cleanCStyle(code);
      if ('Java'==type) code=cleanCStyle(code);
      if ('Javascript'==type) code=cleanCStyle(code);
      if ('C++'==type) code=cleanCStyle(code);
      if ('C#'==type) code=cleanCStyle(code);
      if ('CSS'==type) code=cleanCSS(code);
  }
  else{
      finishTabifier(code);
  }
}

function finishTabifier(code) {
    code=code.replace(/\n\s*\n/g, '\n');  //blank lines
    code=code.replace(/^[\s\n]*/, ''); //leading space
    code=code.replace(/[\s\n]*$/, ''); //trailing space
    
	code=code.replace(/ < /g, '<');
    code=code.replace(/ > /g, '>');

	code=code.replace(/ </g, '<');
    code=code.replace(/ >/g, '>');
	
	code=code.replace(/< /g, '<');
    code=code.replace(/> /g, '>');

    code=code.replace(/</g, ' &lt ');
    code=code.replace(/>/g, ' &gt ');

    //alert("code is formatted, not highlighted");
    code = code + '\n';
    if(highlight === 1){
        if(type=='CSS'){
            code = CSShighlight(code);
        }
        else{
            code = highlightC(code);
        }
    }
    code = "<pre>" + code + "</pre>"
    //document.getElementById("results").outerHTML = code;
    Office.context.document.setSelectedDataAsync(code, { coercionType: 'html' });
    
    //alert("tabifier ends");
    
    level=0;
}

function CSShighlight(code){
    var i = 0;
    var c;
    var out = "";
    
    out = out + "<span style=\"color:blue\">";
    while(i < code.length){
        c = code.charAt(i);
        if(code.substring(i,i+2) === "/*"){
			out = out + "<span style=\"color:green\">";
			while(code.substring(i,i+2) !== "*/"){
				out = out + c;
				i = i + 1;
				c = code.charAt(i);
			}
			out = out + "*/" + "</span>";
			i = i + 2;
			c = code.charAt(i);
		}
        else if(c === ','){
            out = out + "</span>" + ',' + "<span style=\"color:blue\">";
            i = i + 1;
        }
        else if(c === ':'){
            out = out + "</span>" + ':' + "<span style=\"color:orange\">";
            i = i + 1;
        }
        else if(c === '{'){
            out = out + "</span>" + '{' + "<span style=\"color:#585858\">";
            i = i + 1;
            c = code.charAt(i);
            while(c !== '}'){
                if(c === ':'){
                    out = out + "</span>" + ':';
                }
                else if(c === ';'){
                    out = out + ';' + "<span style=\"color:#585858\">";
                }
                else {
                    out = out + c;
                }
                i = i + 1;
                c = code.charAt(i);
            }
            out = out + "</span>" + '}' + "<span style=\"color:blue\">";
            i = i + 1;
        }
        else{
            out = out + c;
            i = i + 1;
        }     
    }
    return out;
}

// returns C Style code highlighted with html
function highlightC(code){
    var i = 0;
    var c;
    var out = "";
	var exp = 1; // set to 1 if expecting a keyword
    while(i < code.length){
        c = code.charAt(i);
		if(code.substring(i,i+2) === "//"){
			out = out + "<span style=\"color:green\">";
			while(c !== '\n'){
				out = out + c;
				i = i + 1;
				c = code.charAt(i);
			}
			out = out + '\n' + "</span>";
			i = i + 1;
			c = code.charAt(i);
			exp = 1;
		}
		else if(code.substring(i,i+2) === "/*"){
			out = out + "<span style=\"color:green\">";
			while(code.substring(i,i+2) !== "*/"){
				out = out + c;
				i = i + 1;
				c = code.charAt(i);
			}
			out = out + "*/" + "</span>";
			i = i + 2;
			c = code.charAt(i);
			exp = 1;
		}
		else if(c === '"'){
			out = out + "<span style=\"color:#585858\">";
			out = out + c;
			i = i + 1;
			c = code.charAt(i);
			while(c !== '"'){
				out = out + c;
				i = i + 1;
				c = code.charAt(i);
			}
			out = out + '"' + "</span>";
			i = i + 1;
			c = code.charAt(i);
			exp = 0;
		}
		else if(c === "'"){
			out = out + "<span style=\"color:#585858\">";
			out = out + c;
			i = i + 1;
			c = code.charAt(i);
			while(c !== "'"){
				out = out + c;
				i = i + 1;
				c = code.charAt(i);
			}
			out = out + "'" + "</span>";
			i = i + 1;
			c = code.charAt(i);
			exp = 0;
		}
		else if(exp === 1){
			if(code.substring(i,i+4) === "int "){
				out = out + "<span style=\"color:blue\">int</span>";
				i = i+3;
			}
			else if(code.substring(i,i+5) === "byte "){
				out = out + "<span style=\"color:blue\">byte</span>";
				i = i+4;
			}
			else if(code.substring(i,i+6) === "short "){
				out = out + "<span style=\"color:blue\">short</span>";
				i = i+5;
			}
			else if(code.substring(i,i+5) === "long "){
				out = out + "<span style=\"color:blue\">long</span>";
				i = i+4;
			}
			else if(code.substring(i,i+6) === "float "){
				out = out + "<span style=\"color:blue\">float</span>";
				i = i+5;
			}
			else if(code.substring(i,i+7) === "double "){
				out = out + "<span style=\"color:blue\">double</span>";
				i = i+6;
			}
			else if(code.substring(i,i+8) === "boolean "){
				out = out + "<span style=\"color:blue\">boolean</span>";
				i = i+7;
			}
			else if(code.substring(i,i+5) === "char "){
				out = out + "<span style=\"color:blue\">char</span>";
				i = i+4;
			}
			else if(code.substring(i,i+5) === "void "){
				out = out + "<span style=\"color:blue\">void</span>";
				i = i+4;
			}
            else if(code.substring(i,i+4) === "var "){
				out = out + "<span style=\"color:blue\">var</span>";
				i = i+3;
			}
            
			else if(code.substring(i,i+7) === "return "){
				out = out + "<span style=\"color:blue\">return</span>";
				i = i+6;
			}
            
			else if(code.substring(i,i+9) === "abstract "){
				out = out + "<span style=\"color:blue\">abstract</span>";
				i = i+8;
			}
            else if(code.substring(i,i+8) === "extends "){
				out = out + "<span style=\"color:blue\">extends</span>";
				i = i+7;
			}
            else if(code.substring(i,i+6) === "super;" || code.substring(i,i+6) === "super(" || code.substring(i,i+6) === "super."){
				out = out + "<span style=\"color:blue\">super</span>";
				i = i+5;
			}
            else if(code.substring(i,i+5) === "this;" || code.substring(i,i+5) === "this(" || code.substring(i,i+5) === "this."){
				out = out + "<span style=\"color:blue\">this</span>";
				i = i+4;
			}
            
            else if(code.substring(i,i+6) === "catch " || code.substring(i,i+6) === "catch("){
				out = out + "<span style=\"color:blue\">catch</span>";
				i = i+5;
			}
            else if(code.substring(i,i+4) === "try " || code.substring(i,i+4) === "try{"){
				out = out + "<span style=\"color:blue\">try</span>";
				i = i+3;
			}
            
            else if(code.substring(i,i+7) === "switch " || code.substring(i,i+7) === "switch("){
				out = out + "<span style=\"color:blue\">switch</span>";
				i = i+6;
			}
			else if(code.substring(i,i+5) === "case "){
				out = out + "<span style=\"color:blue\">case</span>";
				i = i+4;
			}
			else if(code.substring(i,i+6) === "break;"){
				out = out + "<span style=\"color:blue\">break</span>";
				i = i+5;
			}
			
            else if(code.substring(i,i+8) === "package "){
				out = out + "<span style=\"color:blue\">package</span>";
				i = i+7;
			}
			else if(code.substring(i,i+7) === "import "){
				out = out + "<span style=\"color:blue\">import</span>";
				i = i+6;
			}
			else if(code.substring(i,i+8) === "default "){
				out = out + "<span style=\"color:blue\">default</span>";
				i = i+7;
			}
            else if(code.substring(i,i+9) === "#include "){
				out = out + "<span style=\"color:blue\">#include</span>";
				i = i+8;
			}
            
            else if(code.substring(i,i+4) === "new "){
				out = out + "<span style=\"color:blue\">new</span>";
				i = i+3;
			}
            
            else if(code.substring(i,i+9) === "function "){
				out = out + "<span style=\"color:blue\">function</span>";
				i = i+8;
			}
			else if(code.substring(i,i+7) === "public "){
				out = out + "<span style=\"color:blue\">public</span>";
				i = i+6;
			}
			else if(code.substring(i,i+7) === "static "){
				out = out + "<span style=\"color:blue\">static</span>";
				i = i+6;
			}
			else if(code.substring(i,i+8) === "private "){
				out = out + "<span style=\"color:blue\">private</span>";
				i = i+7;
			}
			else if(code.substring(i,i+6) === "System"){
				out = out + "<span style=\"color:purple\">System</span>";
				i = i+6;
			}
			
			else if(code.substring(i,i+5) === "for ("){
				out = out + "<span style=\"color:orange\">for</span>";
				i = i+3;
			}
			else if(code.substring(i,i+5) === "else " || code.substring(i,i+5) === "else{"){
				out = out + "<span style=\"color:orange\">else</span>";
				i = i+4;
			}
			else if(code.substring(i,i+3) === "if " || code.substring(i,i+3) === "if("){
				out = out + "<span style=\"color:orange\">if</span>";
				i = i+2;
			}
			else if(code.substring(i,i+6) === "while " || code.substring(i,i+6) === "while("){
				out = out + "<span style=\"color:orange\">while</span>";
				i = i+5;
			}
			else if(code.substring(i,i+6) === "class "){
				out = out + "<span style=\"color:orange\">class</span>";
				i = i+5;
			}
            else if(code.substring(i,i+7) === "struct "){
				out = out + "<span style=\"color:orange\">struct</span>";
				i = i+6;
			}
			else{
				out = out + c;
				i = i+1;
				exp = 0;
				if(c === ';' || c === ' ' || c === '{' || c === '}' || c === '(' || c === '\n'){
					exp = 1;
				}
			}
		}
		else{
			out = out + c;
			i = i+1;
			if(c === ';' || c === ' ' || c === '{' || c === '}' || c === '(' || c === '\n'){
				exp = 1;
			}
		}
    }
    return out;

}

function repeat(pattern, count) {
    if (count < 1) return '';
    var result = '';
    while (count > 0) {
        if (count & 1) result += pattern;
        count >>= 1, pattern += pattern;
    }
    return result;
}
function tabs() {
  var s='';
  for (var j=0; j<level; j++) s+=repeat(' ', tabstop);
  return s;
}

function cleanCSS(code) {
  var i=0, instring=false, incomment=false, c, cp;
  function cleanAsync() {
    var iStart=i;
    for (; i<code.length && i<iStart+LOOP_SIZE; i++) {
      c=code.charAt(i);
      cp=null;
      try {
        cp=code.charAt(i+1);
      } catch (e) { }

      if (incomment) {
        if ('*' == c && '/' == cp) {
          incomment=false;
          out+='*/';
          i++;
        } else {
          out+=c;
        }
      } else if (instring) {
        if (instring==c) {
          instring=false;
        }
        out+=c;
      } else if ('/'==c && '*'==cp) {
        incomment=true;
        out+='/*';
        i++;
      } else if ('{'==c) {
        level++;
        out+=' {\n'+tabs();
      } else if ('}'==c) {
        out=out.replace(/\s*$/, '');
        level--;
        out+='\n'+tabs()+'}\n'+tabs();
      } else if ('"'==c || "'"==c) {
        if (instring && c==instring) {
          instring=false;
        } else {
          instring=c;
        }
        out+=c;
      } else if (';'==c) {
        out+=';\n'+tabs();
      } else if ('\n'==c) {
        out+='\n'+tabs();
      } else {
        out+=c;
      }
    }

    if (i<code.length) {
      setTimeout(cleanAsync, 0);
    } else {
      level=li;
      out=out.replace(/[\s\n]*$/, '');
      finishTabifier(out);
    }
  }

  if ('\n'==code[0]) code=code.substr(1);
  code=code.replace(/([^\/])?\n*/g, '$1');
  code=code.replace(/\n\s+/g, '\n');
  code=code.replace(/[   ]+/g, ' ');
  code=code.replace(/\s?([;:{},+>])\s?/g, '$1');
  code=code.replace(/\{(.*):(.*)\}/g, '{$1: $2}');

  var out=tabs(), li=level;
  cleanAsync();
  return out;
}




function cleanCStyle(code) {
  var i=0;
  function cleanAsync() {
    var iStart=i;
    for (; i<code.length && i<iStart+LOOP_SIZE; i++) {
      c=code.charAt(i);

      if (incomment) {
        if ('//'==incomment && '\n'==c) {
          incomment=false;
        } else if ('/*'==incomment && '*/'==code.substr(i, 2)) {
          incomment=false;
          c='*/\n';
          i++;
        }
        if (!incomment) {
          while (code.charAt(++i).match(/\s/)) ;; i--;
          c+=tabs();
        }
        out+=c;
      } else if (instring) {
        if (instring==c && // this string closes at the next matching quote
          // unless it was escaped, or the escape is escaped
          ('\\'!=code.charAt(i-1) || '\\'==code.charAt(i-2))
        ) {
          instring=false;
        }
        out+=c;
      } else if (infor && '('==c) {
        infor++;
        out+=c;
      } else if (infor && ')'==c) {
        infor--;
        out+=c;
      } else if ('else'==code.substr(i, 4)) {
        out=out.replace(/\s*$/, '')+' e';
      } else if (code.substr(i).match(/^for\s*\(/)) {
        infor=1;
        out+='for (';
        while ('('!=code.charAt(++i)) ;;
      } else if ('//'==code.substr(i, 2)) {
        incomment='//';
        out+='//';
        i++;
      } else if ('/*'==code.substr(i, 2)) {
        incomment='/*';
        out+='\n'+tabs()+'/*';
        i++;
      } else if ('"'==c || "'"==c) {
        if (instring && c==instring) {
          instring=false;
        } else {
          instring=c;
        }
        out+=c;
      } else if ('{'==c) {
        level++;
        out=out.replace(/\s*$/, '')+' {\n'+tabs();
        while (code.charAt(++i).match(/\s/)) ;; i--;
      } else if ('}'==c) {
        out=out.replace(/\s*$/, '');
        level--;
        out+='\n'+tabs()+'}\n'+tabs();
        while (code.charAt(++i).match(/\s/)) ;; i--;
      } else if (';'==c && !infor) {
        out+=';\n'+tabs();
        while (code.charAt(++i).match(/\s/)) ;; i--;
      } else if ('\n'==c) {
        out+='\n'+tabs();
      } else {
        out+=c;
      }
    }

    if (i<code.length) {
      setTimeout(cleanAsync, 0);
    } else {
      level=li;
      out=out.replace(/[\s\n]*$/, '');
      finishTabifier(out);
    }
  }

  code=code.replace(/^[\s\n]*/, ''); //leading space
  code=code.replace(/[\s\n]*$/, ''); //trailing space
  code=code.replace(/[\n\r]+/g, '\n'); //collapse newlines

  var out=tabs(), li=level, c='';
  var infor=false, forcount=0, instring=false, incomment=false;
  cleanAsync();
}
