          <div class="mod mod-pro-ul">
              <div class="m-body">
                 <div class="m-header">
				   <h3>产品分类</h3>
				   <div class="m-header-ext"><a href="#" class="more">查看所有分类&gt;&gt;</a></div>
				 </div>
                 
                 <div class="m-content m-pro-type" id="m-pro-type-box">
                    <ul class="frist">
					   <li><a href="#">产品分类1</a></li>
					   <li><a href="#">产品分类2</a><em><a href="javascript:void(0);">close</a></em><ul><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li></ul></li>
					   <li><a href="#">产品分类3</a></li>
					   <li><a href="#">产品分类4</a></li>
					</ul>
					<ul class="m-pro-box"></ul>
					<ul>
					   <li><a href="#">产品分类5</a></li>
					   <li><a href="#">产品分类6</a></li>
					   <li><a href="#">产品分类7</a><em><a href="javascript:void(0);">close</a></em><ul><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li></ul></li>
					   <li><a href="#">产品分类8</a></li>
					</ul>
					<ul class="m-pro-box"></ul>
					<ul class="last">
					   <li><a href="#">产品分类9</a></li>
					   <li><a href="#">产品分类10</a><em><a href="javascript:void(0);">&nbsp;</a></em><ul><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li><li><a href="#">产品分类2</a></li></ul></li>
					   <li><a href="#">产品分类11</a></li>
					   <li><a href="#">产品分类12</a></li>
					</ul>
					<ul class="m-pro-box"></ul>
                 </div>
				 
				 <div class="clear">&nbsp;</div>
				 
              </div>
			  
          </div>
		  
		  <script language="javascript">
		  $(function(){
		      //绑定分类展开对象
		      var proULs1 = new proUL($('#m-pro-type-box'));
			  
		      //定义分类展开对象
			  function proUL($ulObj){
			     //打开指定的分类UL
			     this.open = function(){
					 if($onEm!=null){
						$onEm.attr('class','open').html('open');
						$onEm.parent().parent().attr('class','open');
						this.thisUlobj($onEm).css({'border-bottom':'0'});
						this.nextUlopen();
					 }
					 return this;
				 }
			     //关闭指定的分类UL
			     this.close = function(){
					 if($onEm!=null){
						$onEm.attr('class','close').html('close');
						$onEm.parent().parent().attr('class','close');
						this.thisUlobj($onEm).css({'border-bottom':'#ddd 1px solid'});
						this.nextUlclose($onEm);
					 }
					 return this;
				 }
			     //关闭指定的分类UL
			     this.closeOld = function(){
					 if($_onEm!=null){
						$_onEm.attr('class','close').html('close');
						$_onEm.parent().parent().attr('class','close');
						this.thisUlobj($_onEm).css({'border-bottom':'#ddd 1px solid'});
						this.nextUlclose($_onEm);
					 }
					 return this;
				 }
				 //获取当前分类项所在的UL
				 this.thisUlobj = function($obj){
				     if($obj!=null){
						 return $obj.parent().parent().parent();
					 }else{
						 return false;
					 }
				 }
				 //获取相应的子分类
				 this.nextUlobj = function($obj){
				     if($obj!=null){
						 return this.thisUlobj($obj).next();
					 }else{
						 return false;
					 }
				 }
				 //打开相应的子分类
				 this.nextUlopen = function(){
				     if($onEm!=null){
						 return this.nextUlobj($onEm).html($onEmHtml).fadeIn(300);
					 }else{
						 return false;
					 }
				 }
				 //打开相应的子分类
				 this.nextUlclose = function($obj){
				     if($obj!=null){
						 return this.nextUlobj($obj).fadeOut(300);
					 }else{
						 return false;
					 }
				 }


				 /**
				  * ***************************************
				  *  *** 初始化对象.记录当前打开的分类 *******
				  * ***************************************
				  */
				 var $onEm = null;
				 var $_onEm = null;
				 var $onEmHtml = null;
				 var _this = this;
				 /**
				  * ********************
				  *  *** 产品分类逻辑 ***
				  * ********************
				  */
				 var $proLi = $ulObj.find('ul').find('li');
				 $proLi.find('em').find('a').click(function(){
				     //初始化
				     $onEm = $(this);
					 //当前被点击元素.所在的下一个UL
					 var $proNextUl =  _this.nextUlobj( $onEm );
					 //当前被点击元素.附带的二级分类内容(LI)
					 $onEmHtml = $(this).parent().parent().find('ul').html();
					 //各元素符合,则执行
					 if($proNextUl.attr('class')==='m-pro-box' && $onEmHtml!=''){
						 var thisOn = $(this).attr('class');
						 if(thisOn==='open'){
							_this.close();
						 }else{
							_this.closeOld();
							_this.open( $onEmHtml );
							$_onEm = $onEm;
						 }
					 }else{
					     $(this).fadeOut(300);
					 }
				 });
				 //返回对象
				 return this;
			  }

		  });
		  </script>