	if (mtDropDown.isSupported()) {

		var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 0, mtDropDown.reference.bottomLeft);

		var menu1 = ms.addMenu(document.getElementById("menu1"));
		menu1.addItem("- ��˾���","Aboutus.asp?Title=��ҵ���"); 
		menu1.addItem("- �ܲ��´�","Aboutus.asp?Title=�ܲ��´�");
		menu1.addItem("- ��֯����","Aboutus.asp?Title=��֯����");
		menu1.addItem("- �ɳ�����","Aboutus.asp?Title=�ɳ�����");
		menu1.addItem("- ��ϵ����","Aboutus.asp?Title=��ϵ����");		
		menu1.addItem("- ��ҵ�Ļ�","Culture.asp"); 	
	
//�ڶ��˵�
		//var menu2 = ms.addMenu(document.getElementById("menu2"));
		//menu2.addItem("- ��ҵ����", "news.asp"); 
		//menu2.addItem("- ҵ����Ѷ", "yenews.asp"); // send no URL if nothing should happen onclick
	
		// menu : 3
		var menu3 = ms.addMenu(document.getElementById("menu3"));
		menu3.addItem("- ��Ʒչʾ", "Product.asp");
		menu3.addItem("- ��Ʒ����", "Products.asp");	   
		
		// menu : 4
		var menu4 = ms.addMenu(document.getElementById("menu4"));
		menu4.addItem("- ��ҵ����", "CompHonor.asp");
		menu4.addItem("- ��ҵ����", "CompVisualize.asp");
		
		// menu : 5
		var menu5 = ms.addMenu(document.getElementById("menu5"));
		menu5.addItem("- �����г�", "HomeMarket.asp");
		menu5.addItem("- �����г�", "OverseasMarket.asp");

		// menu : 6
		var menu6 = ms.addMenu(document.getElementById("menu6"));
		menu6.addItem("- �˲���Ƹ", "HrDemand.asp");
		menu6.addItem("- �˲Ų���", "HrPolicy.asp");
		
		// menu : 7
		var menu7 = ms.addMenu(document.getElementById("menu7"));
		menu7.addItem("- �ͻ�����", "FeedbackView.asp");
		menu7.addItem("- ��Ҫ����", "Feedback.asp");						
			
		// menu : 8
		var menu8 = ms.addMenu(document.getElementById("menu8"));		
		menu8.addItem("- �޸Ļ�Ա����", "UserEdit.asp");			
		menu8.addItem("- �޸Ļ�Ա����", "UserEditPwd.asp");			
		menu8.addItem("- ��Ʒѯ�۲�ѯ", "UserOrder.asp");
		menu8.addItem("- �鿴�ҵ�����", "FeedbackMember.asp");
		menu8.addItem("- �˳���Ա����", "UserLogout.asp?Action=Ch");
				
	
		mtDropDown.renderAll();
	}
