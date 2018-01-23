var trEE_NODES={
	format:{
		left:18,
		top:170, 
		width:195,
		height:372,
		e_image:"/admin/images/fo_p.gif",		 
		c_image:"/admin/images/fc_p.gif",		 
		i_image:"/admin/images/i_p.gif",
		b_image:'/admin/images/b.gif',
		bgcolor:"#e2e2e2",
		back_bgcolor:"#e2e2e2",
		animation:0,
		padding:2,
		level_ident:16,
		dont_resize_back:1
	},
	sub:[
		
		{html:'Custom Web Pages',
			sub:[
				{html:'Manage Web Pages', url:'/admin/Documents/ManageDocuments.aspx'},
				{html:'Manage Page Categories', url:'/admin/Documents/ManageDocumentCategories.aspx'}
			]
		},
		
				
		{html:'Recipes',
			sub:[
				{html:'Recipes', url:'/admin/Recipes/ManageRecipes.aspx'}				
			]
		
		},
		
		{html:'Manage Users',
			sub:[
				{html:'Manage Access', url:'/Admin/ManageAccess/ManageUsers.aspx'}				
			]
		
		}
		
	]
}
