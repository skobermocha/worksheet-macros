var address = this.getField('client').value + '-' + this.getField('siteaddress').value;
var passing_results = 'On';
var Results_message = 'No Results'
var add_cc = ''

if (this.getField('scope-dt').value == 'Yes') {
	if ((this.getField('dt-results').value == 'NC') && (this.getField('dt-smoke').value !== 'True')) {
		passing_results = 'Off';
		};

};

if (this.getField('scope-af').value == 'Yes') {
	if (this.getField('af-results').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope-fw').value == 'Yes') {
	if (this.getField('fw-results').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope-rcm').value == 'Yes') {
	if ((this.getField('Results_RCM').value == 'NC') || (this.getField('Results_RCM_Weightin').value == 'NC')) {
		passing_results = 'Off';
		};

};

if (passing_results == 'On') {
	Results_message = 'All Pass';}
else {
	Results_message = 'NC Reported';
	add_cc = '&cc=danielleluna@ducttesters.com,cassieh@ducttesters.com,conorpeterson@ducttesters.com';
};

var mailtoUrl = 'mailto:aorresults@ducttesters.com?' + add_cc + '&body=Dont for get to attach the pictures you took to this email.&subject=Rater Results-' + Results_message + ' ' + address ;

this.submitForm({ 
 cURL: mailtoUrl, 
 cSubmitAs: "PDF" 

});
