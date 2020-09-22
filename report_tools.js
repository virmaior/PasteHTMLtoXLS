

function rank_rt(rt_id)
{
	var ranks = $("#rt_" + rt_id ).find(".rank_TD");
	console.log('rt_ ' + rt_id  + ' has ' + ranks.length + ' ranks');
	ranks.each(function(){ 
	  var my_id = $(this).attr('id');
	  var my_rt = my_id.split('_')[0];
	  var y = my_id.split('_')[1];
	  var x = my_id.split('_')[2];
	  $(this).attr('grand_total',$("#" + my_rt + "_" + y + "_" + (x-1)).html());  
//	  console.log("#" + my_rt + "_" + y + "_" + (x-1) +   $("#" + my_rt + "_" + y + "_" + (x-1)).html());
	  });
	sorted_ranks = ranks.sort((a, b) => parseFloat($(b).attr('grand_total')) - parseFloat($(a).attr('grand_total')));
	var rank_count = 1;
	var equals = 1;
	var current_value = 0;
	sorted_ranks.each(function(){
		$(this).html(rank_count);
		if ($(this).attr('grand_total')  > current_value) {
			rank_count += equals;
			equals = 1;
		} else { equals++; }
	});
console.log('rt_ ' + rt_id  + ' has been sorted.'); 
}

function make_column_letter(val)
{
	var result = "";

	while (val > 0) {
		var right_part = val % 26;
		result = String.fromCharCode(right_part + 64) + result;
		val = (val - right_part) / 26 ;
	}
	
	return result;
}


//https://stackoverflow.com/questions/9975707/use-jquery-select-to-select-contents-of-a-div
jQuery.fn.selectText = function(){
    this.find('input').each(function() {
        if($(this).prev().length == 0 || !$(this).prev().hasClass('p_copy')) { 
            $('<p class="p_copy" style="position: absolute; z-index: -1;"></p>').insertBefore($(this));
        }
        $(this).prev().html($(this).val());
    });
    var doc = document;
    var element = this[0];
    console.log(this, element);
    if (doc.body.createTextRange) {
        var range = document.body.createTextRange();
        range.moveToElementText(element);
        range.select();
    } else if (window.getSelection) {
        var selection = window.getSelection();        
        var range = document.createRange();
        range.selectNodeContents(element);
        selection.removeAllRanges();
        selection.addRange(range);
    }
};

$(document).ready(function() {
 $('button.RT_paste_button').on('click',function(){

   var rt_id = $(this).attr('rt_id');
 console.log("copy action fired for " + rt_id);
 	var my_text =$("#rt_" + rt_id).html();
	var simplify = $.parseHTML( '<table>' +  my_text  + '</table>');
	$(simplify).find('.hidden_column').each(function() {$(this).remove()} );
	$(simplify).find('.XLS_switch').each(function() {
		var fmla = $(this).attr('x:fmla');
		var x = $(this).parent().children().index($(this)) + 1 ; //convert to start at 1
		console.log("considering formula " + fmla );
		var run_count =0;
		while (fmla.indexOf(String.fromCharCode(189)) > -1) {
			run_count++;
			var instruction_start = fmla.indexOf(String.fromCharCode(189));
			var instruction_end = fmla.indexOf(String.fromCharCode(190));
			if (instruction_end > -1) {
				instruction_end = instruction_end + 1; //correct to express inclusion
				var instruction = fmla.substring(instruction_start,instruction_end);
				var parts = instruction.split("");
			
			console.log('magic instuction: ' + instruction + ' found. Length = ' + instruction.length + ' (=  ' + (instruction_end - instruction_start ) +  ') command = ' + parts[0] + ' mode = ' + parts[1] + instruction.substr(3,instruction.length -4)); 
			{
				
				if (parts[2] == 'A') {
					var target_id = '#' + instruction.substr(3,instruction.length - 4); 
					var target_x = $(target_id).parent().children().index  ($(target_id)) + 1;
					var xletter = make_column_letter(target_x);
				}
				if (parts[2] == '-') { var xletter = make_column_letter(x  - parts[3]);  } 
				if (parts[2] == '+') { var xletter = make_column_letter(x + parts[3]);  } 
				fmla = fmla.split(instruction).join(xletter);
			} 
				if (run_count > 50) return false;
			}else { 
				console.log("invalid magic formula length");
				return false; 
			}
		}
		console.log("resolved to " + fmla );
		$(this).html(fmla );
		});


	$("#RT_pb_" + rt_id).html(  '<x:Workbook ' +
	' xmlns="urn:schemas-microsoft-com:office:spreadsheet" ' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office" ' + 
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel" ' + 
	'><Worksheet ss:Name="Test"><table >' + $(simplify).html() + '</table></Worksheet></Workbook>');
	 console.log("content set for #RT_pb_" + rt_id  );
	$('#RT_pb_' + rt_id).selectText();
    document.execCommand("copy");
});
});


