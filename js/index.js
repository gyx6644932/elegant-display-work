function zhaolezi(){
  $('.slide-snap-box').on('mouseenter','li',function(e){
    var $this=$(this)
    var index=$this.index();
    $this.siblings('li').removeClass('slide-snap-item-current');
    $this.addClass('slide-snap-item-current');
    $('.slide-box .slide-item').stop(true,true).not(':eq('+index+')').hide();
    $('.slide-box .slide-item').eq(index).show();
    $('.slide-box-masker').stop(true,true).css({'opacity':'1','display':'block'}).fadeOut("fast");
  })
}
zhaolezi()