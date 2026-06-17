import { Log } from '@microsoft/sp-core-library';
import styles from './BpclKrApplicationCustomizer.module.scss';
import Logo from './assets/Bpcl-logo.png';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

const LOG_SOURCE = 'BpclKrApplicationCustomizer';
export default class BpclKrApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _top: PlaceholderContent | undefined;

  public async onInit(): Promise<void> {

    Log.info(
      LOG_SOURCE,
      'Initialized'
    );

    this._hideAppBar();

    await this._renderTop();

    this._renderBottom();

    this.context.application.navigatedEvent.add(
      this,
      () => {
        this._renderBottom();
      }
    );

    return Promise.resolve();
  }

  private _hideAppBar(): void {

    const style = document.createElement(
      'style'
    );

    style.innerHTML = `
       #spSiteHeader { display: none !important; }
      #sp-appBar { display: none !important; }
      .ms-HorizontalNav { margin-left: 0 !important; }
      #spCommandBar { display: none !important; }
      #vpc_Page\\.SiteFooter\\.internal\\.03025612-a400-4804-a78e-e1493200a43b { display: none !important; }
      #CommentsWrapper { display: none !important; }
      
    `;

    document.head.appendChild(style);
  }

private async _renderTop(): Promise<void> {

  if (document.getElementById('bpcl-custom-navbar')) {
  return;
}

  // if (this._top) {
  //   return;
  // }

  this._top =
    this.context.placeholderProvider
      .tryCreateContent(
        PlaceholderName.Top
      );

  if (
    !this._top ||
    !this._top.domElement
  ) {
    return;
  }

this._top.domElement.innerHTML = `
 
<div id="bpcl-custom-navbar" class="${styles.mainNavbar}">
 
  <div class="${styles.navbarContainer}">
 
    <div class="${styles.logoSection}">
<img src="${Logo}" alt="BPCL Logo" class="${styles.logoImg}" />
<span class="${styles.logoText}">KR Connect</span>
</div>
 
    <button
      id="bpclMenuBtn"
      class="${styles.menuBtn}"
      aria-label="Toggle Menu">
      ☰
    </button>
 
    <div
      id="bpclMenu"
      class="${styles.navMenu}">
 
      <div class="${styles.navLinks}">
      <a href="#" class="${styles.navLink} ${styles.active}">HOME</a>
      <a href="#" class="${styles.navLink}">BPMAI</a>
      <a href="#" class="${styles.navLink}">MY PORTAL</a>
      <a href="#" class="${styles.navLink}">FORMS</a>
      <a href="#" class="${styles.navLink}">REFINERY VISION</a>
      <a href="#" class="${styles.navLink}">I-CONNECT</a>
      <a href="#" class="${styles.navLink}">ERDMP</a>
      <a href="#" class="${styles.navLink}">ZOHO REMOTE</a>
      <a href="#" class="${styles.navLink}">AFPM</a>
    </div>
 
    </div>

 
  </div>
 
</div>
 
`;
 

setTimeout(() => {
 
  const menuBtn =
    document.getElementById('bpclMenuBtn');
 
  const menu =
    document.getElementById('bpclMenu');
 
  if (menuBtn && menu) {
 
    menuBtn.addEventListener('click', () => {
 
      const isOpen =
        menu.classList.contains('bpcl-mobile-open');
 
      if (isOpen) {
 
        menu.classList.remove('bpcl-mobile-open');
        menuBtn.textContent = '☰';
 
      } else {
 
        menu.classList.add('bpcl-mobile-open');
        menuBtn.textContent = '✕';
 
      }
 
    });
 
  }
 
}, 300);
 

  }


  private _renderBottom():
    void {

    if (
      document.getElementById(
        'bpcl-footer'
      )
    ) {
      return;
    }

    const pageCanvas =
      document.querySelector(
        '#spPageCanvasContent'
      );

    if (!pageCanvas) {
      return;
    }

    const year =
      new Date()
        .getFullYear();

    const footer =
      document.createElement(
        'div'
      );

    footer.id =
      'bpcl-footer';

    footer.innerHTML = `

 <div class="${styles.footerSection}">

  <div class="${styles.footerContent}">

   <div class="${styles.footerLeft}">
    © ${year}
    Bharat Petroleum Corporation Limited
   </div>
    
    <div class="${styles.footerRight}">
    <a href="#">Feedback</a>
    <a href="#">Help</a>
    </div>

  </div>

</div>

`;

    pageCanvas.appendChild(
      footer
    );

  }

}