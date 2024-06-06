import axios from 'axios';
import Util from './util';

const axiosInstance = axios.create({
  baseURL: Util.BASE_API.dev + '/api/v1.0/plugin/', // url = base url + request url
  headers: {
    'Content-Type': 'application/json'
  },
  timeout: 60 * 60 * 1000
  // timeout: 30 * 1000
  // withCredentials: true, // send cookies when cross-domain requests
})

axiosInstance.interceptors.response.use(
  response => {
    if (response.data.ErrorMsg) {
      throw new Error(response.data.ErrorMsg)
    }
    return response.data
  },
  error => {
    return Promise.reject(error)
  }
)

function translate(text, src, tgt, domain) {
  return axiosInstance.post('translate', 
    {
      src: text,
      srclang: src,
      tgtlang: tgt,
      domain: domain,
      type: 'text'
    }
  );
}

function getLang(text) {
  return axiosInstance.post('get_langid', 
    {
      text: text
    }
  );
}

export default {
  translate,
  getLang
}