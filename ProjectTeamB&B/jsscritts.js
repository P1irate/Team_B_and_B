
$(document).ready(function(){

    $('.nav a').click(function(){

        /*задали какой мы хотим отступ от верха страницы*/
        var otstupTop=100;

        $('body,html').animate({

        /*получили положение элемента вычли отступ и прокрутили*/
           scrollTop: $($(this).attr('href')).offset().top-otstupTop
        }, 1500);

    });
});

