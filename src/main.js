import Vue from 'vue'
import App from './App.vue'
import VueRouter from 'vue-router'
import ribbon from './components/ribbon.js'

Vue.use(VueRouter)
Vue.config.productionTip = false

const routerCfg= [
  {
    path: '/', 
    name: '默认页',
    component:()=>import('./components/HelloWps.vue')
  },{
    path: '/dialog', 
    name: '对话框',
    component:()=>import('./components/Dialog.vue')
  },{
    path: '/taskpane', 
    name: '任务窗格',
    component:()=>import('./components/TaskPane.vue')
  }
]

new Vue({
  render: h => h(App),
  router: new VueRouter({routes:routerCfg}),
  created: function () {
    window.ribbon = ribbon   //把ribbon.js中的函数导出，用于在ribbon.xml中配置，从而wps客户端程序能够调用
  }
}).$mount('#app')
