import { createApp } from 'vue'
import './style.css'
import App from './App.vue'
import router from './router/index'
import{createPinia} from "pinia"
import ElementUI from "element-plus";
import "element-plus/dist/index.css";
const app = createApp(App)
const pinia =createPinia()

app.use(router)
app.use(pinia)
app.use(ElementUI)

app.mount('#app')
