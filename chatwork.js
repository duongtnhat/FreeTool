var request = require('request');

var apiKey = 'API_KEY'

setInterval(function() {
request({
      headers: {
        'X-ChatWorkToken': apiKey
      },
      uri: 'https://api.chatwork.com/v2/my/status'
    }, function (error, response, body) {
    if (!error && response.statusCode == 200) {
	var json = JSON.parse(body);
        console.log(new Date(), 'Unread_room_num:', json.unread_room_num)
	if (json.unread_room_num > 0) {
	  request({
            headers: {'X-ChatWorkToken': apiKey},
            uri: 'https://api.chatwork.com/v2/rooms'
            }, function (error, response, body) {
		var json = JSON.parse(body);
		json.forEach(function(room){
		  if (room.unread_num > 0) console.log("\tRoom:", room.name, " https://www.chatwork.com/#!rid"+room.room_id);
		});
            })
	}
     }
});

}, 60000);
