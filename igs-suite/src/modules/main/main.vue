<template lang="pug">
#app
	.header
		h1 "Hola"
</template>

<script>
export default {
  name: "oneMain"
}
</script>

<style lang="scss">

#app {
    display: grid;
    position: absolute;
    width: 100%;
    height: 100%;
    grid-template-columns: 48px 1fr;
    grid-template-rows: 48px 1fr;
}
.header {
    grid-area: 1 / span 2 / 1 / span 2;

}
</style>
