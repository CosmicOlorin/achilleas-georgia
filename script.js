if('scrollRestoration' in history)history.scrollRestoration='manual';
window.scrollTo(0,0);window.addEventListener('pageshow',()=>window.scrollTo(0,0));window.addEventListener('beforeunload',()=>window.scrollTo(0,0));
const ENDPOINT='https://script.google.com/macros/s/AKfycbyx-z_paY1vRoWfbs9kJOEYdmHj7K0cidVq695U_TOFav64Dj_Vds524-MV3NNwxsCIJg/exec';
const $=s=>document.querySelector(s);const soundtrack=new Audio('beautiful-in-white.mp3?v=36');soundtrack.loop=true;soundtrack.preload='auto';let isPlaying=false;
$('#music').style.zIndex='9999';$('#music').style.pointerEvents='auto';
async function play(){try{await soundtrack.play();isPlaying=true;$('#music').classList.add('playing');$('#music').classList.remove('muted');$('#music').textContent='♫'}catch{isPlaying=false}}
function pause(){soundtrack.pause();isPlaying=false;$('#music').classList.remove('playing');$('#music').classList.add('muted');$('#music').textContent='♪'}
$('#envelope').addEventListener('click',()=>{$('#envelope').classList.add('is-open');play();setTimeout(()=>{window.scrollTo(0,0);$('#intro').classList.add('open');$('#content').classList.add('visible');$('#content').removeAttribute('aria-hidden');document.body.classList.remove('locked');$('#music').classList.add('show');requestAnimationFrame(()=>window.scrollTo(0,0))},1450)},{once:true});
$('#music').addEventListener('click',()=>isPlaying?pause():play());
const wedding=new Date('2026-11-08T13:00:00+02:00');function tick(){const d=Math.max(0,wedding-Date.now()),v=[Math.floor(d/864e5),Math.floor(d/36e5)%24,Math.floor(d/6e4)%60,Math.floor(d/1e3)%60];['days','hours','minutes','seconds'].forEach((id,i)=>$('#'+id).textContent=String(v[i]).padStart(2,'0'))}tick();setInterval(tick,1000);
const visitorId=localStorage.weddingVisitorId||crypto.randomUUID();localStorage.weddingVisitorId=visitorId;
async function send(data){if(ENDPOINT.startsWith('__')){console.info('Pending Google endpoint',data);return}const body=new URLSearchParams({...data,visitorId,page:location.href});await fetch(ENDPOINT,{method:'POST',mode:'no-cors',body})}
$('#rsvpForm').addEventListener('submit',async e=>{e.preventDefault();const s=$('#formStatus'),b=e.target.querySelector('button'),data=Object.fromEntries(new FormData(e.target));b.disabled=true;s.textContent='Αποστολή…';try{await send({action:'rsvp',...data});localStorage.weddingGuestName=data.name.trim();s.textContent='Ευχαριστούμε! Η απάντησή σας καταγράφηκε.';e.target.reset()}catch{s.textContent='Δεν έγινε η αποστολή. Δοκιμάστε ξανά.'}finally{b.disabled=false}});
function legacyCopy(value){const input=document.createElement('textarea');input.value=value;input.setAttribute('readonly','');input.style.cssText='position:fixed;opacity:0;pointer-events:none';document.body.appendChild(input);input.select();const copied=document.execCommand('copy');input.remove();return copied}
document.querySelectorAll('.copy-gift').forEach(button=>button.addEventListener('click',async e=>{
  const current=e.currentTarget,value=current.dataset.copy,gift=current.dataset.gift,status=current.closest('.bank-card').querySelector('.copy-status');
  status.textContent='Γίνεται αντιγραφή…';
  const tracking=send({action:'gift_click',gift,guestName:localStorage.weddingGuestName||''}).catch(()=>{});
  let copied=false;
  try{await navigator.clipboard.writeText(value);copied=true}catch{copied=legacyCopy(value)}
  status.textContent=copied?'✓ Αντιγράφηκε!':'Πατήστε παρατεταμένα για αντιγραφή';
  status.classList.toggle('copied',copied);
  await tracking;
}));
$('#wishForm').addEventListener('submit',async e=>{e.preventDefault();const s=$('#wishStatus');s.textContent='Αποστολή…';try{await send({action:'wish',...Object.fromEntries(new FormData(e.target))});s.textContent='Η ευχή σας στάλθηκε. Ευχαριστούμε!';e.target.reset()}catch{s.textContent='Δεν έγινε η αποστολή. Δοκιμάστε ξανά.'}});
const uploadButton=$('.soft-button'),uploadStatus=$('#uploadStatus');
let uploadWidget;
uploadButton.addEventListener('click',()=>{
  if(!window.cloudinary){uploadStatus.textContent='Η μεταφόρτωση δεν φορτώθηκε. Ανανεώστε τη σελίδα και δοκιμάστε ξανά.';return}
  if(!uploadWidget)uploadWidget=window.cloudinary.createUploadWidget({
    cloudName:'pmmzygig',uploadPreset:'wedding_guests_2026',sources:['local','camera'],multiple:true,
    clientAllowedFormats:['jpg','jpeg','png','webp','heic'],maxFileSize:15000000,maxFiles:20,
    showAdvancedOptions:false,showCompletedButton:true,singleUploadAutoClose:false,
    text:{el:{menu:{files:'Από τη συσκευή',camera:'Κάμερα'},local:{browse:'Επιλογή φωτογραφιών',dd_title_single:'Σύρετε μία φωτογραφία εδώ',dd_title_multi:'Σύρετε φωτογραφίες εδώ'},queue:{title:'Αρχεία για ανέβασμα',title_uploading:'Οι φωτογραφίες ανεβαίνουν',title_uploading_with_counter:'Ανεβαίνουν {{num}} φωτογραφίες',title_uploading_processing:'Γίνεται επεξεργασία',done:'Ολοκληρώθηκε',mini_title:'Ανέβηκαν',mini_title_uploading:'Ανεβαίνουν'},notifications:{general_error:'Κάτι πήγε στραβά. Δοκιμάστε ξανά.',completed:'Η μεταφόρτωση ολοκληρώθηκε.'}}}
  },(error,result)=>{
    if(error){uploadStatus.textContent='Δεν ολοκληρώθηκε η μεταφόρτωση. Δοκιμάστε ξανά.';return}
    if(result&&result.event==='success')uploadStatus.textContent='Η φωτογραφία ανέβηκε. Ευχαριστούμε!';
    if(result&&result.event==='queues-end')uploadStatus.textContent='Οι φωτογραφίες ανέβηκαν. Ευχαριστούμε!';
  });
  uploadWidget.open();
});
const revealObserver=new IntersectionObserver(entries=>entries.forEach(entry=>{if(entry.isIntersecting){entry.target.classList.add('revealed');revealObserver.unobserve(entry.target)}}),{threshold:.14});document.querySelectorAll('.reveal-up,.reveal-down').forEach(el=>revealObserver.observe(el));
const scrollFlowers=[...document.querySelectorAll('.scroll-flower')];let flowerFrame=0;function moveFlowers(){flowerFrame=0;const y=window.scrollY;scrollFlowers.forEach((el,i)=>{const base=Number(el.dataset.base||0),direction=i%2?-1:1,rotation=base+direction*y*.055,scale=1+Math.min(y,1800)/6000;el.style.transform=`rotate(${rotation}deg) scale(${scale})`})}window.addEventListener('scroll',()=>{if(!flowerFrame)flowerFrame=requestAnimationFrame(moveFlowers)},{passive:true});moveFlowers();
