var address = this.getField('subdivision').value + '-Lot ' + this.getField('lotnumber').value + ' ' + this.getField('address').value;
var passing_results = 'On';
var Results_message = 'No Results'
var add_cc = ''

if (this.getField('scope_DT').value == 'Yes') {
	if (this.getField('Results_DT').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_DLTO').value == 'Yes') {
	if (this.getField('Results_DLTO').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_QII').value == 'Yes') {
	if ((this.getField('Results_ENV21H').value == 'NC') || (this.getField('Results_ENV22A').value == 'NC') || (this.getField('Results_ENV21G').value == 'NC') || (this.getField('Results_ENV21F').value == 'NC') || (this.getField('Results_ENV23B').value == 'NC') || (this.getField('Results_ENV23E').value == 'NC') || (this.getField('Results_ENV23C').value == 'NC')){
		passing_results = 'Off';
		};

};
if (this.getField('scope_AF').value == 'Yes') {
	if (this.getField('Results_AF').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_FW').value == 'Yes') {
	if (this.getField('Results_FW').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_RCM').value == 'Yes') {
	if ((this.getField('Results_RCM').value == 'NC') || (this.getField('Results_RCM_Weightin').value == 'NC')) {
		passing_results = 'Off';
		};

};
if (this.getField('scope_BD').value == 'Yes') {
	if (this.getField('Results_BD').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_DD').value == 'Yes') {
	if (this.getField('Results_DD').value == 'NC') {
		passing_results = 'Off';
		};
	if (this.getField('Results_BDucts').value == 'NC') {
		passing_results = 'Off';
		};
	if (this.getField('Results_DBD').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_IAQ').value == 'Yes') {
	if (this.getField('Results_IAQ').value == 'NC') {
		passing_results = 'Off';
		};

};
if (this.getField('scope_DICS').value == 'Yes') {
	if (this.getField('Results_DICS').value == 'NC') {
		passing_results = 'Off';
		};

};

if (passing_results == 'On') {
	Results_message = 'All Pass';}
else {
	Results_message = 'NC Reported';
	add_cc = '&cc=cassieh@ducttesters.com,conorpeterson@ducttesters.com';
};

var mailtoUrl = 'mailto:rncresults@ducttesters.com?' + add_cc + '&body=Dont for get to attach the pictures you took to this email.&subject=Rater Results-' + Results_message + ' ' + address ;

this.submitForm({ 
 cURL: mailtoUrl, 
 cSubmitAs: "PDF" 

});
