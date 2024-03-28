import { defineStore } from "pinia";

//定义仓库
export const useNum=defineStore('firstStoreId',{
    state:()=>{
        return{
            name:'张三',
            age:18,
        }
    },
    //计算属性
    getters:{
        doubleAge(){
            return this.age*2
        },
    },
    //函数
    actions:{
        //同步方法
        changeAge(val){
            this.age+=val
        },
        //异步
        asyncChangeAge(){
            return new Promise((resolve, reject) => {
                setTimeout(()=>{
                    this.age+=1
                    resolve()
                },1000)
            })
        },
        aaa(){
            nihao(x=>2*x)
        }
    },

})