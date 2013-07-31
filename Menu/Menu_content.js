	if (mtDropDown.isSupported()) {

		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 0, mtDropDown.reference.bottomLeft);

		var menu1 = ms.addMenu(document.getElementById("menu1"));
		menu1.addItem("- 公司简介","Aboutus.asp?Title=企业简介"); 
		menu1.addItem("- 总裁致辞","Aboutus.asp?Title=总裁致辞");
		menu1.addItem("- 组织机构","Aboutus.asp?Title=组织机构");
		menu1.addItem("- 成长历程","Aboutus.asp?Title=成长历程");
		menu1.addItem("- 联系我们","Aboutus.asp?Title=联系我们");		
		menu1.addItem("- 企业文化","Culture.asp"); 	
	
//第二菜单
		//var menu2 = ms.addMenu(document.getElementById("menu2"));
		//menu2.addItem("- 企业新闻", "news.asp"); 
		//menu2.addItem("- 业内资讯", "yenews.asp"); // send no URL if nothing should happen onclick
	
		// menu : 3
		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem("- 产品展示", "Product.asp");
		menu3.addItem("- 产品分类", "Products.asp");	   
		
		// menu : 4
		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem("- 企业荣誉", "CompHonor.asp");
		menu4.addItem("- 企业形象", "CompVisualize.asp");
		
		// menu : 5
		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem("- 国内市场", "HomeMarket.asp");
		menu5.addItem("- 国外市场", "OverseasMarket.asp");

		// menu : 6
		var menu6 = ms.addMenu(document.getElementById("menu6"));
		menu6.addItem("- 人才招聘", "HrDemand.asp");
		menu6.addItem("- 人才策略", "HrPolicy.asp");
		
		// menu : 7
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		menu7.addItem("- 客户留言", "FeedbackView.asp");
		menu7.addItem("- 我要留言", "Feedback.asp");						
			
		// menu : 8
		var menu8 = ms.addMenu(document.getElementById("menu8"));		
		menu8.addItem("- 修改会员资料", "UserEdit.asp");			
		menu8.addItem("- 修改会员密码", "UserEditPwd.asp");			
		menu8.addItem("- 产品询价查询", "UserOrder.asp");
		menu8.addItem("- 查看我的留言", "FeedbackMember.asp");
		menu8.addItem("- 退出会员中心", "UserLogout.asp?Action=Ch");
				
	
		mtDropDown.renderAll();
	}
