// ==================================================
// 出退勤チェックを判定せずに空回しする
// ==================================================
(function(global){
  var ParupunteManager = (function() {
    
    // ==================================================
    // コンストラクタ
    // 実行タイプ別に変数をセットする
    // ==================================================
    function ParupunteManager(type){
      this.processType = type;
      switch(type){
        // from all to all
        case 0:
          this.AccountToken = ACCOUNT_TOKEN_PARUPUNTA;
          this.inputRoomID = ROOM_ID_ALL;
          this.replyRoomID = ROOM_ID_ALL;
          break;
        // from all to person
        case 1:
          this.AccountToken = ACCOUNT_TOKEN_PARUPUNTA;
          this.inputRoomID = ROOM_ID_ALL;
          this.replyRoomID = 0;
          break;
        // from person to person
        case 2:
          this.AccountToken = ACCOUNT_TOKEN_PARUPUNTA;
          this.inputRoomID = 0;
          this.replyRoomID = 0;
          break;
      }
    };
    
    // ==================================================
    // 入力の読み込みとログへの書き込み
    // ==================================================
    ParupunteManager.prototype.processGet = function(){
      try{
        switch(this.processType){
            // 0 : from all(from all to all)
          case 0:
            _processGet(this.AccountToken,this.inputRoomID);
            break;
            // 1 : from all(from all to person)
          case 1:
            _processGet(this.AccountToken,this.inputRoomID);
            break;
            // 2 : from person(from peson to person)
          case 2:
            _processGet(this.AccountToken,'');
            break;
        }
      }catch(e){
        new Parupuntetest.ChatWork({token: this.AccountToken}).sendMessage(
          {room_id: ROOM_ID_MANAGER, body: MESSAGE_LIST['err001']+e.message}
        );
      }
    };
    
    // ==================================================
    // 入力漏れのチェック
    // ==================================================
    ParupunteManager.prototype.processCheckInputMiss = function(){
      try{
        switch(this.processType){
            // 0 : from all(from all to all)
          case 0:
            _checkInputMiss(this.AccountToken);
            break;
            // 1 : from all(from all to person)
          case 1:
            _checkInputMiss(this.AccountToken);
            break;
            // 2 : from person(from peson to person)
          case 2:
            _checkInputMiss(this.AccountToken);
            break;
        }
      }catch(e){
        new Parupuntetest.ChatWork({token: this.AccountToken}).sendMessage(
          {room_id: ROOM_ID_MANAGER, body: MESSAGE_LIST['err001']+e.message}
        );
      }
    };
    
    // ==================================================
    // 返信
    // ==================================================
    ParupunteManager.prototype.processSet = function(){
      try{
        _processSet(this.AccountToken,this.processType);
      }catch(e){
        new Parupuntetest.ChatWork({token: this.AccountToken}).sendMessage(
          {room_id: ROOM_ID_MANAGER, body: MESSAGE_LIST['err001']+e.message}
        );
      }
    };
      
    return ParupunteManager;
  })();

  global.ParupunteManager = ParupunteManager;
})(this);