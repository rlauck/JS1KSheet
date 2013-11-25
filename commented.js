var global=this,
	ls=localStorage,
	inputs={}; // map of col+row ref string => cell input element - this leaks mem due to the circular reference in the event handler closure
function ref(v,i,c,r){
	// i,c,r are placeholders. Pass a 0-based index to get an Excel-style cell reference or pass a cell reference string to get a 1-based index
	for(r=i=v>=0?ref(v/26-1)+(10+v%26|0).toString(36):'';c=v[i++];r=r*26+parseInt(c,36)-9); // upper-case only alternative for (10+v%26|0).toString(36): String.fromCharCode(65+v%26)
	return r;
}
function calculate(id){ 
	//TODO: cache vals, set an "in-progress" val (#REF?) before the eval in global[id]() and use to detect cycles instead of try/catch
	for(id in inputs) // recalculate every cell
		try { inputs[id].value = global[id]() || [ls[id]]; } catch(x) {}
}
function addRow(row,j,rowNode){ // row is a 0-based row index, j and rowNode are placeholders
	rowNode=global.T.insertRow(row), // window.[element id] might not work in every browser, but its so much shorter than document.getElementById('T')
		ls.y=row+1; // store the total table length relying on rows being added in order
	function addCell(col, node, id){ // col is the column reference (ex: 'AA'), node is the row html element, id is 0 for row and column headers, nonzero otherwise 
		node.innerHTML = id ? '<input size=10 style=border:0>' : '<center>'+(row||col); // style=text-align:center is waaay too many bytes - W3C are clearly not golfers
		if(id)
			inputs[id=col+row]=node=node.firstChild, // repurpose id as the cell reference (ex: 'A1'), node is set to the new input element 
			node.onblur = function(){
				calculate(ls[id]=node.value); // persist the cell value in localStorage then recalculate all cells
			},
			node.onfocus = function(){
				node.value=[ls[id]]; // load the cell value from storage for editing
			},
			node.onkeyup = function(e){ //TODO: it'd be cool to support arrow keys
				e.which-13 // short-circuit unless the Enter key (13) is pressed
					|| (ls.y > (e=row+1) // reuse e while checking if this is not the last row
						|| !addRow(e))  // if it is the last row, add a new one - no return value so this will evaluate to true
					&& inputs[col+e].focus(); // the row should now exist, so focus on the cell in the next row and same column
			},
			global[id] = function(v){ // store the cell value in the global context (function if it's a formula, number/string/undefined otherwise)
				with(Math) // add Math to the scope of eval
				return /^=/.test(v=ls[id]||0) // test for formulas while also retreiving the cell value from localStorage in the placeholder v
					? +eval(v.slice(1).replace(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/gi,function(q,c1,r1,c2,r2){ // strip '=', replace cell refs with function calls and eval the formula
						q=[], // repurpose q to hold an array of cell ref strings
						r2=r2?r2-r1+1:1; // repurpose r2 to hold the # of rows in the range
						// the following loops through the range implied by a cell ref/range.
						// examples of regex match => output:
						//   'A1' => 'A1()'
						//   'A1:B2' => 'B2(),A2(),B1(),A1()' 
						// c2 is reused as the loop counter and c1 holds the index of the first column
						for(c2=r2*(ref(c2||c1)-(c1=ref(c1)-1));c2--;) // Pseudocode: c1=startColumn; c2=numRows * numColumns; while(c2-- > 0)...
							q[c2]=ref(c1+~~c2/r2)+(+r1+c2%r2)+'()'; // Pseudocode: c1+~~c2/r2 => startColumn + Math.floor(counter / numRows), +r1+c2%r2 => parseInt(startRow) + (counter % numRows)
						return''+q; // the default behavior when the array is coerced to string is equivalent to q.join(',')
					})) 
					: v==+v?+v:v // coerce v if it's a number or falsy, leave as a string otherwise. undefined is coerced to 0 to avoid NaN errors in formulas that include empty cells
			}
	}
	for(j=-1;j<99;) // you can change the constant 99 here to adjust the # of columns
		addCell(ref(j), rowNode.insertCell(++j), j*row);
}
function sum(){
	// Array.reduce was that shortest way I found to implement sum for variable arguments
	// If you don't pass a second argument, it treats the first element as the initial value and starts with the second
	return [].reduce.call(arguments,function(p,c){return p+c}); 
}
function avg(){
	// this should be obvious. 0 is passed as the first arg of apply since we never rely on the context in sum
	return sum.apply(0,arguments)/arguments.length;
}
for(var i=0, len=ls.y||9; i<len; addRow(i++)); // localStorage.y holds the table length, or default to 9 rows initially
calculate(); // this could be left out, but then no values are loaded from localStorage until the first onblur call
