var address = this.getField('subdivision').value + '-Lot ' + this.getField('lotnumber').value + ' ' + this.getField('address').value;
var passing_results = 'On';
var Results_message = 'No Results'
var add_cc = ''


if ((this.getField('Results_ENV21A').value == 'NC') || (this.getField('Results_ENV21B').value == 'NC') || (this.getField('Results_ENV21C').value == 'NC') || (this.getField('Results_ENV21D').value == 'NC') || (this.getField('Results_ENV21E').value == 'NC') || (this.getField('Results_ENV21F').value == 'NC') || (this.getField('Results_ENV21G').value == 'NC') || (this.getField('Results_ENV21H').value == 'NC')){
	passing_results = 'Off';
};

if ((this.getField('Results_ENV22A').value == 'NC') || (this.getField('Results_ENV22B').value == 'NC') || (this.getField('Results_ENV23A').value == 'NC') || (this.getField('Results_ENV23B').value == 'NC') || (this.getField('Results_ENV23C').value == 'NC') || (this.getField('Results_ENV23D').value == 'NC') || (this.getField('Results_ENV23E').value == 'NC') || (this.getField('Results_ENV23F').value == 'NC') || (this.getField('Results_ENV23G').value == 'NC') || (this.getField('Results_ENV23H').value == 'NC')){
	passing_results = 'Off';
};
if ((this.getField('Results_PLB22F').value == 'NC') || (this.getField('Results_PLB22G').value == 'NC') || (this.getField('Results_PLB22H').value == 'NC') || (this.getField('Results_PLB22I').value == 'NC') || (this.getField('Results_PLB22J').value == 'NC')){
	passing_results = 'Off';
};

if (this.getField('Insul_Matching').value == 'No') {
	passing_results = 'Off';
};
if (this.getField('Results_MC').value == 'NC') {
	passing_results = 'Off';
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
