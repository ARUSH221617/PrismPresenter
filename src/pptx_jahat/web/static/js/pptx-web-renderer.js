/**
 * pptx-web-renderer.js
 * Client-side Web PPTX Vector & Interactive DOM Renderer.
 */

class PPTXWebClientRenderer {
  constructor(containerId) {
    this.container = document.getElementById(containerId);
    this.slidesHtml = [];
    this.currentIdx = 0;
  }

  async loadFromHtmlApi(filePath, width = 800) {
    try {
      const res = await fetch('/api/preview/slide-html', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_path: filePath, width })
      });
      const data = await res.json();
      if (data.success && data.slides && data.slides.length > 0) {
        this.slidesHtml = data.slides;
        this.currentIdx = 0;
        this.renderCurrent();
        return { success: true, count: this.slidesHtml.length, engine: 'Web Render Engine' };
      }
      return { success: false, error: data.error || 'No slides returned' };
    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  renderCurrent() {
    if (!this.container || !this.slidesHtml || this.slidesHtml.length === 0) {
      return;
    }
    const htmlContent = this.slidesHtml[this.currentIdx];
    this.container.innerHTML = htmlContent;
  }

  nextSlide() {
    if (this.currentIdx < this.slidesHtml.length - 1) {
      this.currentIdx++;
      this.renderCurrent();
      return this.currentIdx;
    }
    return this.currentIdx;
  }

  prevSlide() {
    if (this.currentIdx > 0) {
      this.currentIdx--;
      this.renderCurrent();
      return this.currentIdx;
    }
    return this.currentIdx;
  }
}

window.PPTXWebClientRenderer = PPTXWebClientRenderer;
