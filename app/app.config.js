(function () {
	var tenantName = 'contoso';
	
	var officeAddin = angular.module('officeAddin');
	officeAddin.constant('appId', '00000000-0000-0000-0000-000000000000');
	officeAddin.constant('sharePointUrl', 'https://' + tenantName + '.sharepoint.com');
})();