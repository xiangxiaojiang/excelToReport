<script setup >
import { ref, onMounted } from "vue";
// import { nihao } from "./typescript/aaa";

// 响应式状态
const count = ref(0);

// 用来修改状态、触发更新的函数
function increment() {
  count.value++;
}

// 生命周期钩子
onMounted(() => {
  // nihao();
});
</script>

<template>
  <router-view />
</template>


<style scoped>
</style>
