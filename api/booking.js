const GAS_URL = process.env.GAS_URL;

function json(res, status, body) {
  res.status(status).setHeader('Content-Type', 'application/json; charset=utf-8').send(JSON.stringify(body));
}

export default async function handler(req, res) {
  if (!GAS_URL) return json(res, 500, { success:false, message:'Environment variable GAS_URL belum dikonfigurasi di Vercel.' });
  try {
    const target = new URL(GAS_URL);
    const params = new URLSearchParams(req.query || {});
    if (req.method === 'GET') {
      params.forEach((value,key)=>target.searchParams.set(key,value));
      const r = await fetch(target.toString(), { method:'GET', headers:{Accept:'application/json'} });
      const text = await r.text();
      res.status(r.status).setHeader('Content-Type','application/json; charset=utf-8').send(text);
      return;
    }
    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
      const r = await fetch(target.toString(), { method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8',Accept:'application/json'}, body });
      const text = await r.text();
      res.status(r.status).setHeader('Content-Type','application/json; charset=utf-8').send(text);
      return;
    }
    res.setHeader('Allow','GET, POST');
    return json(res,405,{success:false,message:'Method tidak didukung.'});
  } catch (err) {
    return json(res,502,{success:false,message:err.message || 'Gagal menghubungi Google Apps Script.'});
  }
}
