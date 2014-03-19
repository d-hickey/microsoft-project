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
var tabstop;
var lang;

function runTabifier() {
  var code = document.getElementById('code').value;
  var type=document.getElementById('mydropdown');
  type=type.options[type.selectedIndex].value;

  lang = type;
  
  tabstop=document.getElementById('spacepicker');
  tabstop=tabstop.options[tabstop.selectedIndex].value;

  //console.log(tabstop+"ok");

  if ('C'==type) code=cleanCStyle(code);
  if ('Java'==type) code=cleanCStyle(code);
  if ('Javascript'==type) code=cleanCStyle(code);
  if ('C++'==type) code=cleanCStyle(code);
  if ('C#'==type) code=cleanCStyle(code);
  if ('CSS'==type) code=cleanCSS(code);
}

function finishTabifier(code) {
  code=code.replace(/\n\s*\n/g, '\n');  //blank lines
  code=code.replace(/^[\s\n]*/, ''); //leading space
  code=code.replace(/[\s\n]*$/, ''); //trailing space
  
  // makes get request to syntax highlighting api
  $.get(
    "http://markup.su/api/highlighter",
    {language : lang, theme : 'IDLE', source : code},
    function(data) {
        document.getElementById("results").outerHTML = data;
        Office.context.document.setSelectedDataAsync(data, { coercionType: 'html' });
    }
  );

  level=0;
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
