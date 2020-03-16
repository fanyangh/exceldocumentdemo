import Vue from 'vue'
import Router from 'vue-router'
import doc from '@/components/doc'
import doctemp from '@/components/doctemp'

Vue.use(Router)

export default new Router({
  routes: [
    {
      path: '/',
      name: 'doc',
      component: doc
    },
    {
      path: '/doc',
      name: 'doca',
      component: doctemp
    }

  ]
})
