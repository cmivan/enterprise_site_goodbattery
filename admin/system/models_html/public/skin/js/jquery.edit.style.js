//��������ƹ���ʽ
$(function(){
	
	$(".type_box").find(".typeitem").hover(
	   function(){$(this).css({"background-color":"#fff","border":"#F90 1px solid"});},
	   function(){$(this).css({"background-color":"","border":"#ccc 1px solid"});}
	);
		
	
//������Ϣɾ��	
  $("a.del_item").click(function(){
	 var thisid =$(this).attr("id");
	 var thisnum=$(this).parent().parent().find("span.num").text();
	 var thisok=true;
	 var thisstr="ȷ��Ҫɾ������Ŀ/ҳ����";
	 if(parseInt(thisnum)==thisnum){
	    if(thisnum>0){
		   thisok=false;
		}else{
		   thisok=true;	
		}
	 }else{
		thisok=true;	
	 }
	 
	 if(thisok==false){thisstr = "����Ŀ�»��� "+thisnum+" ����¼��"+thisstr;}
	 
	 if(confirm(thisstr)){
	   var thislink=$(this).attr("link");
	   $(this).attr("href",thislink);
	   return true;		
	 }else{
	   return false; 
	 }
	});
	 

//ȫѡ��ȡ���б���
  $("#delsel").click(function(){
	 var thischeck=$(this).attr("checked");
	 $(".delitem").attr("checked",thischeck);
	 });
	 
  $("#del_checkbox").click(function(){
	 var thischeck=$(this).attr("checked");
	 $(".del_id").attr("checked",thischeck);
	 });

  $("#Submit_delsel").click(function(){
	  if(ischecked()){
		if(confirm("ȷ��Ҫɾ��ѡ����?")){
			return true;
		}else{
			return false;	
		}
	  }else{
		return false;
	  }
	});
	


	
});


/*�ж��Ƿ�ѡ����*/
function ischecked(){
	var thisis=false;
	$(".del_id").each(function(){
		if(thisis==false){
			var thischecked=$(this).attr("checked");
			if(thischecked){thisis=true;}
			}
		});
	if(thisis==false){alert("����Ҫѡ��һ��!");}
	return thisis;
	}