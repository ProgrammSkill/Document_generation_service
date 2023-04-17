//

var i = 1;
$("#btn_add_row").on("click", function(e){
     i++;
     $('#table tbody').append('<tr>' +
    '<td><input type="text" class="form-control" name="col-1_row' + i + '"></td>' +
    '<td><input type="text" class="form-control" name="col-2_row' + i + '"></td>' +
    '<td><input type="number" class="form-control" name="col-3_row' + i + '"></td>')
 });

//var name_service;
//var type_service;
//var price;
//var count = 0;
//$("#btn_add_row").on("click", function(e){
//    name_service = $("#form_document input[name='name_service']").val();
//    type_service = $("#form_document input[name='type_service']").val();
//    price = $("#form_document input[name='price']").val();
//    count = count + 1;
//    $('#table tbody').append('<tr>' +
//    '<td name="col1_row_' + count +'">' + name_service + '</td>' +
//    '<td name="col2_row_' + count +'">' + type_service + '</td>' +
//    '<td name="col3_row_' + count +'">' + price + '</td></tr>')
//});


$("#person_proxy").on("click", function(e){
    $("#block_person_proxy").css('display','block');
    $('.required_input').attr('required', true);
});
$("#director").on("click", function(e){
    $("#block_person_proxy").css('display','none');
    $('.required_input').attr('required', false);
});


$(".document").on("click", function(e){
    $("#alien_data_block").css('display','block');
    $('.required_input').attr('required', true);
});
$("#absent").on("click", function(e){
    $("#alien_data_block").css('display','none');
    $('.required_input').attr('required', false);
});


$("#individual").on("click", function(e){
    $("#individual_data_block").css('display','block');
    $('.required_input_individual').attr('required', true);
});
$("#legal_entity").on("click", function(e){
    $("#individual_data_block").css('display','none');
    $('.required_input_individual').attr('required', false);
});

