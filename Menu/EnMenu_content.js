	if (mtDropDown.isSupported()) {

		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 0, mtDropDown.reference.bottomLeft);

		var menu1 = ms.addMenu(document.getElementById("menu1"));
		menu1.addItem("- About us","EnAboutus.asp?Title=About us"); 
		menu1.addItem("- President oration","EnAboutus.asp?Title=President oration");
		menu1.addItem("- Structure","EnAboutus.asp?Title=Structure");
		menu1.addItem("- History","EnAboutus.asp?Title=History");
		menu1.addItem("- Contact us","EnAboutus.asp?Title=Contact us");			
	
//第二菜单
		//var menu2 = ms.addMenu(document.getElementById("menu2"));
		//menu2.addItem("- 企业新闻", "news.asp"); 
		//menu2.addItem("- 业内资讯", "yenews.asp"); // send no URL if nothing should happen onclick
	
		// menu : 3
		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem("- Product list", "EnProduct.asp");
		menu3.addItem("- Product class", "EnProducts.asp");	   
		
		// menu : 4
		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem("- Honor", "EnCompHonor.asp");
		menu4.addItem("- Visualize", "EnCompVisualize.asp");
		
		// menu : 5
		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem("- HomeMarket", "EnHomeMarket.asp");
		menu5.addItem("- OverseasMarket", "EnOverseasMarket.asp");

		// menu : 6
		//var menu6 = ms.addMenu(document.getElementById("menu6"));
		//menu6.addItem("- 人才招聘", "HrDemand.asp");
		//menu6.addItem("- 人才策略", "HrPolicy.asp");
		
		// menu : 7
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		menu7.addItem("- FeedbackView", "EnFeedbackView.asp");
		menu7.addItem("- Add Feedback", "EnFeedback.asp");						
			
		// menu : 8
		var menu8 = ms.addMenu(document.getElementById("menu8"));		
		menu8.addItem("- Modify Member's Data", "EnUserEdit.asp");			
		menu8.addItem("- Modify Member's Password", "EnUserEditPwd.asp");			
		menu8.addItem("- Product Inquire Search", "EnUserOrder.asp");
		menu8.addItem("- My Feedback", "EnFeedbackMember.asp");
		menu8.addItem("- Member Logout", "UserLogout.asp?Action=En");
				
	
		mtDropDown.renderAll();
	}
