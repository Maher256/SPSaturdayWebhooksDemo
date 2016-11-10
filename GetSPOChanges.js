var csomapi = require('csom-node');
var settings = require('./config');

csomapi.setLoaderOptions({url: settings.url});  

var authCtx = new AuthenticationContext(settings.url);
var changetoken = null;
var changeTypes = ["No Change", "Add", "Update", "Delete", "Rename"]

function _getChanges(io, relativeURL, listId){

authCtx.acquireTokenForUser(settings.username, settings.password, function (err, data) {
    
    console.log("Connecting to site:" + relativeURL);
    var ctx = new SP.ClientContext(relativeURL);  //set root web
    authCtx.setAuthenticationCookie(ctx);  //authenticate
     
    //retrieve SP.Web client object
    var web = ctx.get_web();
    var list = web.get_lists().getById(listId);
    var cQuery = new SP.ChangeQuery();
    cQuery.set_item(true);
    cQuery.set_add(true);
    cQuery.set_update(true);
    cQuery.set_deleteObject(true);
    cQuery.set_changeTokenStart(changetoken);

    var changes = list.getChanges(cQuery);
    ctx.load(web);
    ctx.load(list);
    ctx.load(changes);
    ctx.executeQueryAsync(function () {
        console.log("Connected to site: " + web.get_title() + ' and found ' + changes.get_count() + ' items.');
        var changeEnumerator = changes.getEnumerator();
        
        while (changeEnumerator.moveNext()) {
            var changeItem = changeEnumerator.get_current();
            var changeType = changeTypes[changeItem.get_changeType()];

            io.emit('news', { 
                filename: changeItem.get_itemId() , 
                image: changeType, 
                description:  changeItem.get_time().toLocaleDateString("en-us"),
                url: settings.url + relativeURL + '/_api//web/lists(guid\''+listId+'\')/items/getById('+ changeItem.get_itemId()+')' }); 

            //set change token
            changetoken = changeItem.get_changeToken();
        }
        
    },
    function (sender, args) {
        console.log('An error occured: ' + args.get_message());
        io.emit('news', { filename: 'Error connecting to SharePoint', image: 'Error', description: args.get_message()  });
    });
      
});

}

module.exports = {
  getChanges: _getChanges
}