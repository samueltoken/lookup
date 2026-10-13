function normalizeWebSourceUrl(value) {
  const text = String(value || '').trim();
  if (!text || text.length > 8192) throw new Error('웹주소 확인 필요');
  let url;
  try { url = new URL(/^[a-z][a-z0-9+.-]*:/i.test(text) ? text : `https://${text}`); }
  catch { throw new Error('웹주소 확인 필요'); }
  if (!['https:', 'http:'].includes(url.protocol) || !url.hostname || url.username || url.password) {
    throw new Error('http 또는 https 주소 입력');
  }
  return url.href;
}

function getWebPdfName(title, sourceUrl) {
  const value = String(title || new URL(sourceUrl).hostname).trim();
  return (value.replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/[.\s]+$/g, '').slice(0, 100) || '웹페이지') + '.pdf';
}

function getPublicProductDetailUrl(sourceUrl) {
  const url = new URL(sourceUrl);
  if (url.hostname.toLowerCase() !== 'item.gmarket.co.kr' || url.pathname.toLowerCase() !== '/item') return '';
  const goodsCode = [...url.searchParams].find(([key]) => key.toLowerCase() === 'goodscode')?.[1];
  if (!/^\d{5,14}$/.test(goodsCode || '')) return '';
  return `https://item.gmarket.co.kr/Item/ItemDetailV2?goodsCode=${goodsCode}&detailVersionType=2&isNew=true&fasterDeliveryYn=false&fromBottomLayer=true`;
}

module.exports = { normalizeWebSourceUrl, getWebPdfName, getPublicProductDetailUrl };
