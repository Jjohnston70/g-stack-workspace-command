var LoggerModule = {

  log: function(message, meta) {
    console.log(JSON.stringify({
      message: message,
      meta: meta || {},
      ts: new Date().toISOString()
    }));
  }

};