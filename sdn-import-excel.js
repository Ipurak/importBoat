//Author : Ham
//Created : Dec 06, 2017 1:43 pm
/*########## ########## ########## ########## ##########   ##########*/
/*########## [START]APP CONVERT TO JSON AND MAPPING COULMN ##########*/
/*########## ########## ########## ########## ##########   ##########*/

var XLStoJSON = (function(){

  var data = {
    url:'',
    element:'',
    workbook:[],
    importSelected:'',
    sheetSelected:'',
    import:''
  }

  var init = function( options ){
    data.element = $( options.el )
    data.url     = options.url
    data.import  = options.import

    //element = my_file_input
    //url     = api.asp

    let GetAjax = Ajax({//[Ajax]Get init import from API 
      import:data.import,
      type:'getinit'
    })

    GetAjax.success(function(res){//[Success]Get init import from API 
      let Json = JSON.parse(res)
      $.each(Json,function(key,obj){
        if(  obj != null ){
          data[key] = JSON.parse(obj)
        }
      })
    })

    GetAjax.error(function(res){//[Error]Get init import from API
      alert("Error!, Get init import from API was fail.")
    })

    data.element.change(function(e){
      filePicked( e )
    })

  }

  var mappingColumn = function () {

    let template = {
      title   :$( `<h4 class="modal-title">Map Columns</h4>` ),
      closeBtnTop:$( `<button type="button" class="close" data-dismiss="modal"><span>&times;</span></button>` ),
      closeBtnButtom:$(`<button type="button" data-dismiss="modal" class="btn btn-primary" style="margin-bottom: 0px;">
                          <i class="fa fa-times"></i> Close
                      </button>`),
      NextBtn:$(`<button type="button" class="btn btn-primary next step1" style="margin-bottom: 0px;" disabled>
                    <i class="fa fa-chevron-right"></i> Next
                </button>`),
      BackBtn:$(`<button type="button" class="btn btn-primary back" style="margin-bottom: 0px;">
                  <i class="fa fa-chevron-left"></i> Back
              </button>`),
      ConfirmBtn:$(`<button type="button" class="btn btn-primary confirm" style="margin-bottom: 0px;">
                  <i class="fa fa-check"></i> Confirm
                </button>`),
      CancelBtn:$(`<button type="button" class="btn btn-primary cancel" style="margin-bottom: 0px;">
                  <i class="fa fa-times"></i> Cancel
                </button>`)
    }

    let testBtn = $('<a>TestLink</a>')
    let modal = BootstrapModal()
        modal.header.append( template.closeBtnTop, template.title )
        modal.body.append( menu() ) 
        modal.footer.append( template.closeBtnButtom, template.NextBtn )
        modal.open()

        template.ConfirmBtn.click(function(){
          SendImport()
        })

        template.NextBtn.click(function(){

          if( $( this ).hasClass( "step1" ) ){//select Sheet to import

            $( this ).removeClass( "step1" )
            $( this ).addClass( "step2" ).attr( "disabled","" )
            $( '.content-step1' ).hide()
            template.closeBtnButtom.hide()
            modal.footer.append( template.BackBtn, template.NextBtn )
            $( '.content-step2' ).remove()
            modal.body.append( sheet() )
            
          }else if( $( this ).hasClass( "step2" ) ){//select sheet's columns to map with sedna's column
            updateProgress.calculate()
            $( this ).removeClass( "step2" )
            $( this ).addClass( "step3" )
            $( '.content-step2' ).hide()
            $( '.content-step3' ).remove()
            modal.body.append( mapBody() )
          }else if( $( this ).hasClass( "step3" ) ){//confirm to import
            $( this ).removeClass( "step3" )
            $( this ).addClass( "step4" )
            $( '.content-step3' ).hide()
            $( '.modal-footer > .back' ).hide()
            $( '.modal-footer > .next' ).hide()
            $( '.content-step4' ).remove()
            modal.body.append( confirmImport() )
            modal.footer.append( template.CancelBtn.show(), template.ConfirmBtn.show() )
          }

        })

        template.BackBtn.click(function(){

          if( $( template.NextBtn ).hasClass( "step2" ) ){//Back to step 1 Select Type to import
            $( template.NextBtn ).removeClass( "step2" )
            $( template.NextBtn ).addClass( "step1" )
            $( '.content-step1' ).show()
            $( '.content-step2' ).hide()
          }else if( $( template.NextBtn ).hasClass( "step3" ) ){//Back to step 2 Select Sheet to import
            $( template.NextBtn ).removeClass( "step3" )
            $( template.NextBtn ).addClass( "step2" )
            $( '.content-step2' ).show()
            $( '.content-step3' ).hide()
          }

        })

        template.CancelBtn.click(function(){//Back to Step 3 Map column
          $( template.NextBtn ).removeClass( "step4" )
          $( template.NextBtn ).addClass( "step3" )
          $( '.content-step3' ).show()
          $( '.content-step4' ).hide()
          template.NextBtn.show()
          template.BackBtn.show()
          template.CancelBtn.hide()
          template.ConfirmBtn.hide()
        })
  }

  var SendImport = function(){
    // console.log( "sheetSelected: ",data.sheetSelected )
    let SheetJson   = JSON.parse( to_json( data.workbook ) )
    let MappingJson = data[data.importSelected]
    let ImportJson  = []
    $.each(SheetJson[data.sheetSelected], function(index,array){
      let tempObj = {}
      $.each(MappingJson.data,function(key,obj){
        tempObj[key] = array[obj.map]
      })
      ImportJson.push(tempObj)
    })
    console.log( "Resault: ",ImportJson )
    Ajax({
      json:JSON.stringify( ImportJson ),
      type:'import'
    })
  }

  var menu = function(){
    let template = {
      row:$( `<div class="row content-step1"></div>` ).css({"height":"470px","overflow":"auto"}),
      col:`<div class="col-md-6"></div>`
    }

    let cssMenu = {
      "width": "100%",
      "padding": "5px",
      "background-color": "#7799CC",
      "margin-top": "5px",
      "line-height": "146px",
      "text-align": "center",
      "color": "#ffffff",
      "cursor": "pointer",
      "white-space": "nowrap",
      "text-overflow": "ellipsis",
      "overflow": "hidden"
    }
    
    $.each(data.import.split(","), function(index,menu){
      if( data[menu] != undefined ){
        let active = ""
        if(index === 0){
          active = "active"
          // data.importSelected = menu
        }
        let col = $( template.col ).append( $('<div class="importType" key="'+menu+'" name="'+data[menu].name+'" >'+'<i class="'+data[menu].icon+' fa-3x"></i>'+" "+data[menu].name+'</div>').css(cssMenu) ).addClass( active )
        $( template.row ).append( col )
      }
    })

    $('.importType', template.row).hover(function() {
      $( this ).addClass("pulse animated").css({"outline":"2px solid rgb(119, 153, 204)","background-color":"#ffffff","color":"rgb(119, 153, 204)"})//hover
    }, function() {
      if( !$(this).hasClass("selected") ){
        $( this ).removeClass("pulse animated").css({"outline":"","background-color":"rgb(119, 153, 204)","color":"#ffffff"})//mouse out
      }else{
        $( this ).css({"color":"#F90"})//mouse out
      }
    })

    $( '.importType', template.row ).click(function(){
      console.log( "Before: ", data.importSelected )
      data.importSelected = $(this).attr("key")//set selected import
      console.log( "After: ", data.importSelected )
      $('.step1').removeAttr("disabled")
      $('.importType').css({"color":"#ffffff","background-color":"rgb(119, 153, 204)"}).removeClass("selected")
      $(this).css({"color":"#F90","background-color":"#ffffff"}).addClass("selected")

    })

    return template.row
  }

  var sheet = function(){
    let template = {
      row:$( `<div class="row content-step2"></div>` ).css({"height":"470px","overflow":"auto"}),
      col:`<div class="col-md-6"></div>`
    }

    let cssMenu = {
      "width": "100%",
      "padding": "5px",
      "background-color": "#7799CC",
      "margin-top": "5px",
      "line-height": "146px",
      "text-align": "center",
      "color": "#ffffff",
      "cursor": "pointer",
      "white-space": "nowrap",
      "text-overflow": "ellipsis",
      "overflow": "hidden"
    }
    console.log( "data.workbook.SheetNames: ",data.workbook.SheetNames )
    console.log( ">>: ",data.workbook )
    $.each(data.workbook.SheetNames, function(index,menu){
      let active = ""
      if(index === 0){
        active = "active"
      }
      let col = $( template.col ).append( $('<div class="importSheet" key="'+menu+'" name="'+menu+'" >'+menu+'</div>').css(cssMenu) ).addClass( active )
      $( template.row ).append( col )
    })

    $('.importSheet', template.row).hover(function() {
      $( this ).addClass("pulse animated")//hover
    }, function() {
      $( this ).removeClass("pulse animated")//mouse out
    })

    $( '.importSheet', template.row ).click(function(){
      $('.importSheet').css("color","#ffffff")
      $(this).css("color","#F90")
      console.log( ">>: ",data.workbook.Sheets[$(this).attr("key")] )
      data.sheetSelected = $(this).attr("key")
      $('.step2').removeAttr("disabled")
    })

    return template.row
  }

  var mapBody = function(  ){

    let template = {
      body:$( `<div class="content-step3"></div>` ),
      content:$( `<div></div>` ).css({ "height":"470px", "overflow":"auto" }),
      importNumber:$(`
        <div class="alert alert-info animated flash">
          <i class="fa fa-info-circle" aria-hidden="true"></i> 
            We found <b>[ <span class="found"></span> ]</b> to import, please match the column from your import with our data.
        </div>
      `),
      hasheader:$(`
        <div class="btn-group">
          <button type="button" class="btn">
            <i class="fa fa-question-circle" aria-hidden="true"></i> Do your import has headers title?
          </button>
          <button type="button" class="btn btn-primary header active" value="1">
            <i class="fa fa-check" aria-hidden="true"></i> Yes
          </button>
          <button type="button" class="btn btn-primary header" value="0">
            <i class="fa fa-times" aria-hidden="true"></i> No
          </button>
        </div>
      `).css({ "margin-top": '5px' }),
      progress:$(`
        <div class="row">
          <div class="col-md-12">
            <div class="progress">
              <div class="progress-bar progress-bar-success" style="width: 0%">0%</div>
              <div class="progress-bar progress-bar-danger"  style="width: 100%">100%</div>
            </div>
          </div>
        </div>
      `).css({ "margin-top": '5px' }),
      select : `
      <div class="form-group">
        <label>Select list:</label>
        <select class="form-control" ></select>
      </div>`
    }

    $.fn.map = function () {
      $( this ).addClass("selected").css({"border-color":"green"})
      updateProgress.selectElements = $( 'select', template.body )
      updateProgress.calculate()
    }

    $.fn.notMap = function () {
      $( this ).removeClass("selected").css({"border-color":"red"})
      updateProgress.selectElements = $( 'select', template.body )
      updateProgress.calculate()
    }

    /*######### [START]Create Select List ##########*/
    $( '.found', template.importNumber ).html( JSON.parse( to_json( data.workbook ) )[data.sheetSelected].length - data[data.importSelected].headerRange )
    let Option           = CreateOption
        Option.range     = data[data.importSelected].headerRange
        Option.sheetName = Object.keys( data.workbook.Sheets )[0]
        OptionElement    = Option.get()

    $.each( data[data.importSelected].data, function( key, object ){
      let selectContent = $( template.select )
          $( 'label', selectContent ).text( object.name )
          $( 'select', selectContent ).attr("key",key).empty().append( OptionElement ).notMap()
          template.content.append( selectContent )

    })
    /*######### [END]Create Select List ##########*/

    /*########## [START]Event Click Has Header? ##########*/
    $( 'button.header', template.hasheader ).click(function(){

      $( 'button.header' ).removeAttr( 'disabled' ).removeClass( 'active' )//active class button clicked
      $( this ).attr( 'disabled', '' ).addClass( 'active' )//disable button clicked
      ResetMapping.reset([data.importSelected])//reset all mapping to null

      let type    = parseInt( $( this ).val() )
      $( '.found', template.importNumber ).html( (JSON.parse( to_json( data.workbook ) )[data.sheetSelected].length) - ((type) ? data[data.importSelected].headerRange : 0) )
      let select  = $( 'select', template.body )
                    $( 'select', template.body ).notMap()
      let Option           = CreateOption
          Option.range     = (type) ? data[data.importSelected].headerRange : 0
          Option.sheetName = Object.keys( data.workbook.Sheets )[0]
          OptionElement    = Option.get()

          $.each( select, function( index, element ){
            $( element ).empty().hide().append( OptionElement ).fadeIn(1000)
          })
      
    })
    /*########## [END]Event Click Has Header? ##########*/


    /*########## [START]Event When Map Select Change ##########*/
    $( 'select', template.content ).change(function(){//Select list Event on Change

      let indexColumnSelected = $( this ).val()
      let KeyMap = $( this ).attr("key")
      
      console.log("indexColumnSelected: ", indexColumnSelected)
      console.log("KeyMap: ", KeyMap)

      if ( $( this ).val() != -1 ){//Selected

        $( this ).map()
        data[data.importSelected].data[KeyMap].map = indexColumnSelected
        console.log("data[data.importSelected].data[KeyMap].map: ", data[data.importSelected].data[KeyMap].map)
        console.log("data: ", data)

      }else{//Not Selected

        $( this ).notMap()
        data[data.importSelected].data[KeyMap].map = null

      }

    })
    /*########## [END]Event When Map Select Change ##########*/

    return $( template.body ).append( template.importNumber, template.hasheader, template.progress, template.content )
    
  }

  var confirmImport = function(){

    var template = {
      body:$('<div class="content-step4"></div>'),
      message:$(`<div class="alert alert-warning">
                  <span class="glyphicon glyphicon-exclamation-sign" aria-hidden="true"></span>
                  Do you want to import?
              </div>`)
    }



    return template.body.append( template.message )

  }

  var updateProgress = {
    selectElements:[],
    calculate:function(){

      let select  = this.selectElements 
      let total  = select.length
      let mapping = select.filter(function() {return $(this).hasClass('selected')})
          mapping = mapping.length
      let notMapping = total - mapping
      let percentMapping = Math.round( (mapping/total)*100 )
      let percentNotMapping = Math.round( (notMapping/total)*100 )
      $( '.progress-bar-success',data.progress ).css( "width",percentMapping+"%" ).html( percentMapping+'%' )
      $( '.progress-bar-danger',data.progress ).css( "width",percentNotMapping+"%" ).html( percentNotMapping+'%' )
      if( mapping === 0 ){
        $('.modal-footer > .next').attr("disabled","")
      }else if( mapping === 1 ){
        $('.modal-footer > .next').removeAttr("disabled")
      }

    }
  }

  var ResetMapping = {
    reset:function(key){
      $.each(data[key].data, function(key,object){
        object.map = null
      })
    }
  }

  var CreateOption = {
    range:0,
    sheetName:"",
    get:function(){
      
      let JSONArr    = JSON.parse( to_json( data.workbook ) )
      let SheetArr   = JSONArr[this.sheetName]
      let optionStr  = `<option></option>`
      let result     = '<option value="-1">--- please choose one ---</option>'
      let range      = this.range
      let sheetName  = this.sheetName
      $.each(JSONArr[this.sheetName][this.range], function(index,value){

        let text = value
        if( range > 0 ){//Has header

          if (Object.keys(data[data.importSelected].data)[index] != undefined){
            let position = data[data.importSelected].data[Object.keys(data[data.importSelected].data)[index]].position
            text = SheetArr[range-1][index]+' ---> '+text.replace("{{empty}}", "empty value at column["+position+"]")
          }
          
        }

        result = result+$( optionStr ).attr({"value":index}).text(text).prop('outerHTML')

      })

      return result
    }
  }

  var filePicked = function ( oEvent ) {

    // Get The File From The Input
    var oFile = oEvent.target.files[0]
    var sFilename = oFile.name
    var typeFile = sFilename.split('.');
        typeFile = typeFile[typeFile.length-1]
    
    if( typeFile === "xls" || typeFile === "xlsx" ){

      var reader = new FileReader()
      // Ready The Event For When A File Gets Selected
      let vm = data
      reader.onload = (function(e) {
        var data = e.target.result
        var cfb = XLSX.read( data, {"type":"binary", WTF:false} )
         vm.workbook = cfb
         mappingColumn( )
      })
      // Tell JS To Start Reading The File.. You could delay this if desired
      reader.readAsBinaryString(oFile)
    }else{
      $('<span>Wrong type</span>').afterInsert(data.element)
      alert("Wrong type of this file")
    }

  }

  var to_json = function (workbook) {

    var result = {}
    workbook.SheetNames.forEach(function(sheetName) {
      // var roa = XLS.utils.sheet_to_json(workbook.Sheets[sheetName], {header:1, defval:"{{empty}}"})
      var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {header:1, defval:"{{empty}}"})
      if(roa.length) result[sheetName] = roa
    })
    return JSON.stringify(result, 2, 2)

  }

  var Ajax = function ( obj ) {

    return $.ajax({
      url:"api.asp",
      data:obj,
      type:"POST"
    })

  }

  return { init:init }

})()

/*########## ########## ########## ########## ########## ##########*/
/*########## [END]APP CONVERT TO JSON AND MAPPING COULMN ##########*/
/*########## ########## ########## ########## ########## ##########*/

/*########## ########## ########## ##########   ##########*/
/*########## [START]Lazy To Use Bootstrap Modal ##########*/
/*########## ########## ########## ##########   ##########*/

var BootstrapModal = function() {

  var Modal = {

    main: $("<div>").attr({class:"modal fade",role:"dialog"}),
    dialog: $("<div>").attr("class","modal-dialog"),
    content: $("<div>").attr("class","modal-content"),
    header: $("<div>").attr("class","modal-header"),
    body: $("<div>").attr("class","modal-body"),
    footer: $("<div>").attr("class","modal-footer"),
    open:function(){
      Modal.main.append(Modal.dialog,Modal.content)
      Modal.content.append(Modal.header,Modal.body,Modal.footer)
      Modal.dialog.append(Modal.content)
      $(Modal.main).modal('show')
      
      $(Modal.main).modal('show').promise().done(function(){
        Modal.hidden()
        Modal.shown()
      })
    },
    hide:function(e){
      return e
    },
    show:function(e){
      return e
    },
    hidden:function(hide){
      if (typeof hide == "function"){
        $(Modal.main).on('hidden.bs.modal', function () {
          $(this).data('bs.modal', null).remove()
          Modal.hide(hide())
        })
      }
    },
    shown:function(show){
      if (typeof show == "function"){
        $(Modal.main).on('shown.bs.modal', function () {
          Modal.show(show())
        })
      }
    }
      
  }//end modal


    return Modal
}

/*########## ########## ########## ##########   ##########*/
/*########## [END]Lazy To Use Bootstrap Modal ##########*/
/*########## ########## ########## ##########   ##########*/

