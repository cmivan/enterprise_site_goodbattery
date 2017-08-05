<!--#include file="mod/header.asp"-->
<body>

<div class="mainBody">

  <!--头部内容-->
  <!--#include file="mod/top.asp"-->

  <!--主体内容-->
  <div class="pageBody">
     <div class="pageLeft">

        <!--供应商信息-->
		<!--#include file="mod/side/supplier.asp"-->
     
        <!--旺铺评价-->
		<!--#include file="mod/side/comment.asp"-->
        
        <!--产品搜索-->
        <!--#include file="mod/side/search.asp"-->
        
        <!--产品分类-->
        <!--#include file="mod/side/pro-type.asp"-->

        <!--热销排行榜-->
        <!--#include file="mod/side/pro-hot.asp"-->

        <!--证书荣誉-->
		<!--#include file="mod/side/certificate.asp"-->

        <!--联系方式-->
		<!--#include file="mod/side/contact.asp"-->
        
        <!--二维码-->
		<!--#include file="mod/side/other.asp"-->
        
        
     </div>
	 
     <div class="pageRight">
        <div class="pageRightBox">
		
		
          <!--产品分类-->
		  <!--#include file="mod/pro-type-box.asp"-->
		  
        
          <!--产品分类-->
          <div class="mod mod-pro-ul">
              <div class="m-body">
                 <div class="m-header">
				 <h3>产品分类</h3>
				 <div class="m-header-ext"><a href="#" class="more">查看所有分类&gt;&gt;</a></div>
				 </div>
                 
                 <div class="m-content">
                    <div class="app-offerGeneral">
                      <ul class="offer-list-row">
					  <%for i=0 to 30%>
                          <li>
                          <div class="image">
                          <a href="" title="供应热销无汞无铅环保扣式电池AG0  LR63/179/521" target="_blank">
                          <img src="public/images/829387268_1648406864.120x120.jpg" alt="供应热销无汞无铅环保扣式电池AG09/521">
                          </a>
                          </div>
                          <div class="price"><span class="cny">￥</span><em>3.00</em></div>
                          <div class="title"><a href="#" title="供应热销无汞无铅环保扣式电池LR63/179/521" target="_blank">供应热销无汞无铅环保扣式电池..</a></div>
                          <div class="attributes"><span class="alipay-icon">&nbsp;</span></div>
                          <div class="booked-count"></div>
                          </li>
					  <%next%>
                      </ul>
                      <div class="clear">&nbsp;</div>
                    </div>
                 </div>
              </div>
          </div>
          
        
        </div>
     </div>
     <div class="clear"></div>
  </div>
  
  <!--底部版权信息-->
  <!--#include file="mod/footer.asp"-->
  
</div>

</body>
</html>