var lastColor = '';

function lnkIn(src,txt)
{
	if (!src.contains(event.fromElement))
	{
  	src.style.cursor = 'hand';
		lastColor = src.bgColor;
  	src.bgColor = '#6699FF';
	}
  window.status = txt;
}

function lnkOut(src)
{
  if (!src.contains(event.toElement))
  {
		src.style.cursor = 'default';
  	src.bgColor = lastColor;
  }
	window.status = '';
}
