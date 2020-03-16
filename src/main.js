import Vue from 'vue'
import App from './App'
import router from './router'

import Vuex from 'vuex'
import 'element-ui/lib/theme-chalk/index.css'
import element from 'element-ui'
Vue.use(Vuex)
Vue.use(element)
Vue.config.productionTip = false
new Vue({
  el: '#app',
  router,
  render: h => h(App)
})
