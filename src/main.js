import Vue from 'vue'
import App from './App.vue'
import router from './router'
import $ from 'jquery'
import VueEvents from 'vue-events'
import Vuex from 'vuex'
import * as VeeValidate from 'vee-validate'

Vue.use($)
Vue.use(VueEvents)
Vue.use(VeeValidate)
Vue.use(Vuex)

Vue.config.productionTip = false

new Vue({
  router,
  render: h => h(App),
}).$mount('#app')
