document.addEventListener('DOMContentLoaded', function(){
  var whatsappNumber = '34613346180';
  var fab = document.getElementById('rr-whatsapp-fab');
  var panel = document.getElementById('rr-whatsapp-panel');
  var closeBtn = document.getElementById('rr-whatsapp-close');

  if(!fab || !panel) return;

  function setPanel(open){
    panel.classList.toggle('open', open);
    panel.setAttribute('aria-hidden', String(!open));
  }

  fab.addEventListener('click', function(){
    setPanel(!panel.classList.contains('open'));
  });

  if(closeBtn){
    closeBtn.addEventListener('click', function(){
      setPanel(false);
    });
  }

  panel.querySelectorAll('[data-wa-text]').forEach(function(button){
    button.addEventListener('click', function(){
      var text = button.getAttribute('data-wa-text') || '';
      window.open('https://wa.me/' + whatsappNumber + '?text=' + encodeURIComponent(text), '_blank', 'noopener');
      setPanel(false);
    });
  });

  document.addEventListener('click', function(event){
    if(!panel.contains(event.target) && !fab.contains(event.target)){
      setPanel(false);
    }
  });
});
