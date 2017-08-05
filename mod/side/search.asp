<%if is_side_search=true then%>

<div class="mod">
<div class="m-body">
<div class="m-header"><h3>产品搜索</h3></div>
<div class="m-content">
<form method="get" action="/page/offerlist.htm">
<div class="part-searchInSite"><dl><dt>产品名</dt>
<dd><input type="text" class="search-keywords" name="keywords" maxlength="50" value=""></dd>
<dt>价格</dt><dd><input type="text" class="price-low" name="priceStart">到<input type="text" class="price-high" name="priceEnd"></dd></dl>
<input type="submit" value="搜索" class="search-btn" data-tracelog="wp_widget_search_mainsearch">
<input type="hidden" name="showType" value="">
<input type="hidden" name="sortType" value="showcase">
</div>   
</form>
</div>
</div>
</div>

<%end if%>