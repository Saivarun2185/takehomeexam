var express = require('express')
var router = express.Router()
var multer = require('multer');
var xlsxj = require("xlsx-to-json");
var _ = require('underscore');
var json2xls = require('json2xls');
var fs = require('fs');
var path = require('path');
var xl = require('excel4node');


var storage = multer.diskStorage({

  destination: function (req, file, cb) {
    cb(null, 'uploads')
  },

  filename: function (req, file, cb) {
    cb(null, 'illendula_input.xlsx')
  }
})
 
var upload = multer({ storage: storage })

/* GET */
router.get('/', function(req, res, next) {
  res.render('index', { success: false })
})

/* POST */
router.post('/', upload.single('inputfile'),(req,res) => {
  
  xlsxj({
    input: "uploads/illendula_input.xlsx",
    output: null
  }, 
  
  function(err, result) {
    if(err) {
      console.error(err)
    }else {
      console.log(result)

    var sorteddata =  _.sortBy( result, function( item ) { return -item['Critic Score'] && item['Genre'] } )

    var border =  { 
      left: {
          style: 'thin', 
          color: '000000' 
      },
      right: {
          style: 'thin',
          color: '000000'
      },
      top: {
          style: 'thin',
          color: '000000'
      },
      bottom: {
          style: 'thin',
          color: '000000'
      }
    }
    
    var wb = new xl.Workbook({
       defaultFont: {
      size: 11,
      name: 'Calibri',
      color: 'FF000000'
      }
     })
    
    
     var ws = wb.addWorksheet('Output');     
   
     var myStyle = wb.createStyle({
    
      font: {
           bold: true,
           color: '000000'
       },
       fill: {
             type: 'pattern',
             patternType: 'solid',
             fgColor: 'C6E0B4' 
         },
         border: border
   })

        ws.cell(2,2,2,4).string('Name').style(myStyle)

        var style1 = wb.createStyle({
          font:{
            underline: true,
            italics: true,
            bold: false
          },
          border: border
        })
       
        ws.cell(2,3,2,4,true).string('Illendula,Saivarun').style(style1)
        
        var style2 = wb.createStyle({
          font:{
            bold: true,
            color: 'FFFFFF'
          },
           fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: 'C00000' 
        },
        border: border,
        alignment:{
          horizontal: 'center'
        }
        })

        var i = 1;
        _.each(['SNO','Genre','Credit Score','Album Name','Artist','Release Date'], function(ele){
          ws.cell(4,i).string(ele).style(style2)
          i++;
        })


        i = 5;
      
        var color1 = 'FFF2CC'
        var color2 = 'C6E0B4'
        var prevgenre = sorteddata[0]['Genre']

     _.each(sorteddata, function(ele){
        ele['Credit Score'] = ele['Critic Score']
        delete ele['Critic Score']
        j = 1;
        if(prevgenre != ele['Genre']){
          temp = color1;
          color1 = color2;
          color2 = temp;
          prevgenre = ele['Genre']
        }
        _.each(['SNO','Genre','Credit Score','Album Name','Artist','Release Date'], function(attr){
          var bodystyle = wb.createStyle({
            fill: {
             type: 'pattern',
             patternType: 'solid',
             fgColor: color1 
            },
            border: border,
            alignment: {
              horizontal: /\d/.test(ele[attr]) ? 'right' : 'left'
            }
         })
          ws.cell(i,j).string(ele[attr]).style(bodystyle)
          j++;
        })
        i++;
        
      })
        wb.write('./public/illendula_output.xlsx', function(err,stats){
          if(err) return res.send('Error')
            res.download(path.join(__dirname, '../public/illendula_output.xlsx'),'illendula_output.xlsx')
        });
        //  res.render('index',{
        //    success: true
        //  })
    }
  })
})
module.exports = router
