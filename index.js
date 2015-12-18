var Excel = require('exceljs');
stream = require('stream');


//Globals
numFlights = 4
var ranks = ["WO2s","FSgts","Sgts","FCpls","Cpls","LACs","ACs"]

//Main
extract("MASTER.xlsx")
.then(transform)
.then(split)
.then(prep)
.then(flight)



function extract(filePath){
	return new Promise(function(resolve,reject){
		console.log("EXTRACT")
		//Cadets
		var cadets = [];
		var workbook = new Excel.Workbook();
		workbook.xlsx.readFile("MASTER.xlsx")
		    .then(function(data) {
		    	for (var i = 1; i < 249; i++) {
		    		var row = data._worksheets[1]._rows[i];
		    		var rank = row._cells[0]._value.model.value;
		    		var lastName =  row._cells[1]._value.model.value; 
		    		var firstName =  row._cells[2]._value.model.value; 
		    		var gender =  row._cells[3]._value.model.value; 
		    		var object = new Object({"rank":rank,"lastName":lastName,"firstName":firstName,"gender":gender})
		    		//console.log(object);
		    		cadets.push(object);	
		    	};
		    	resolve(cadets);
		       
		    });
	})
}

function transform(cadets){
	return new Promise(function(resolve,reject){
		console.log("TRANSFORM")
		for (var i = 0; i < cadets.length; i++) {
			//Provide the cadets rank Numbers for easy sorting later
			var rankNumber = 0;

			var cadet = cadets[i]
			switch(cadet.rank){
				case "WO2": 
					rankNumber = 7; 
					cadet.rankNumber = rankNumber;
				break;
				case "FSgt": 
					rankNumber = 6; 
					cadet.rankNumber = rankNumber;
				break;
				case "Sgt": 
					rankNumber = 5; 
					cadet.rankNumber = rankNumber;
				break;
				case "FCpl": 
					rankNumber = 4; 
					cadet.rankNumber = rankNumber;
				break;
				case "Cpl": 
					rankNumber = 3; 
					cadet.rankNumber = rankNumber;
				break;
				case "LAC": 
					rankNumber = 2; 
					cadet.rankNumber = rankNumber;
				break;
				case "AC": 
					rankNumber = 1; 
					cadet.rankNumber = rankNumber;
				break;
			}
		};
		resolve(cadets);
	})
}

//Split all the cadets up into ranks and genders
function split(cadets){
	return new Promise(function(resolve,reject){
		console.log("SPLIT")
		//TODO: Move this block to a different function
		var squadron = new Object();
		

		//Create the squadron data structure
		for (var i = 0; i < ranks.length; i++) {
			Object.defineProperty(squadron,ranks[i],{
				value: new Object({males:[],females:[]}),
				writable: true,
				enumerable: true,
				configurable: true
			})
		};

		console.log(squadron)

		//For each cadet
		for (var i = 0; i < cadets.length; i++) {
			var cadet = cadets[i];
		//If their rank matches one of the ranks.
			for (var rank in squadron){
				//Take the trailing s out of  the rank (e.g. "FSgt|s")
				var rankTitle = rank.substring(0, rank.length - 1);
				if (cadet.rank == rankTitle){
					if (cadet.gender == "M"){
						squadron[rank].males.push(cadet);
					}
					else{
						squadron[rank].females.push(cadet);
					}
				}

			}
		}
	resolve(squadron);
	})	
}


function prep(squadron){
	return new Promise(function(resolve,reject){
		console.log("PREP")
		//Define the flight object
		var flightObject = {name:"name",
					  cadets:[],
					  getNumRankGender: function(rank,gender){
					  	for (var i = 0; i < this.cadets.length; i++) {
					  		var counter = 0;
					  		var cadet = this.cadets[i];
					  		if(cadet.rank == rank && cadet.gender == gender){
					  			counter ++;
					  		}
					  		return counter;
					  	};
					  },
					  getTotal: function(){
					  	return this.cadets.length;
					  },
					  getNumGender: function(gender){
					  	var counter = 0;
					  		for (var i = 0; i < this.cadets.length; i++) {
					  			var counter = 0;
					  			var cadet = this.cadets[i];
					  			if(cadet.gender == gender){
					  				counter ++;
					  			}
					  		};
				  		return counter;
					  },
					  getStats: function(){
					  	//Get Total
					  	stats = {};
					  	stats.name = this.name;
					  	stats.total = this.getTotal();
					  	stats.males = this.getNumGender("M");
					  	stats.females = this.getNumGender("F");
					  	return stats;
					  }

					 }

 		//Add flights to the squadron
 		console.log("Flights array added")
 		squadron.flights = [];
 		console.log(squadron)



		
		//Create the flights
		for (var i = 1; i <= numFlights; i++) {
			var flight = Object.create(flightObject);
			flight.name = String(i);
			squadron.flights.push(flight)
		};
		
		console.log("Flights added");
		//console.log(squadron)
		resolve(squadron)
	})
}

function flight(squadron){
	return new Promise(function(resolve,reject){
		//Start with the AC's
		console.log("Starting with ACs");
		var flightingRank = "AC"
		var cadets = squadron[flightingRank + "s"].males
	
		//FOR EACH CADET
		for (var i = 0; i < 3; i++) {
			console.log(cadets[i])
			var cadet = cadets[i];	
			//GET THE STATS
			var stats = [];
			for (var flight in squadron.flights){
				//console.log(squadron.flights[flight]);
				stats.push(squadron.flights[flight].getStats());
			}
			console.log(stats);

			//Now that we have the stats, we need to apply our rules. First find the max total
			var maxTotal = 0
			for (var x = 0; x < stats.length; x++) {
				if(stats[x].total > maxTotal) {
					maxTotal = stats[x].total;
				}
			};
			console.log("Max total is " + maxTotal);

			//Case 1: If all we care about is totals, add the cadet to the first empty element in the array. Deterine candidates
			var candidates = []
			for (var y = 0; y < stats.length; y++) {
				if(stats[y].total < maxTotal || stats[y].total == 0){
					candidates.push(stats[y].name)
				}
			};
			console.log("Candidates to push into are " + candidates)
			var candidate = candidates[Math.floor(Math.random()*candidates.length)];
			console.log("Candidate picked is " + candidate)
			squadron.flights[candidate - 1].cadets.push(cadet);
			console.log(squadron.flights[candidate-1])
			console.log("Cadet " + cadet.lastName + ", " +cadet.firstName + " "  + "Pushed into flight " + candidate);
			for (var flight in squadron.flights){
				console.log(squadron.flights[flight])
			}

		};
	})
}
