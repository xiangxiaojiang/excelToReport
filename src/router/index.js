import { createWebHistory, createRouter } from 'vue-router'
import Layout from "@/layout/index.vue";

const routes = [
    {
        path: "/",
        name: "index",
        meta: { title: "首页", icon: "icon-shouye1", menuCode: "0" },
        redirect: "/index/",
        component: Layout,
        children: [
          {
            path: "/index",
            name: "index",
            meta: { title: "首页", role: [] },
            component: () => import("@/views/index.vue"),
          },
        ],
      }
];

const router = createRouter({
    history: createWebHistory(),
    routes: routes,
  });

export default router;
