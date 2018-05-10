import Vue from 'vue'

import OneIconsMenu from "../../components/icons/menu.vue"
import OneIconsMenuVertical from "../../components/icons/menu-vertical.vue"
import OneIconsHome from "../../components/icons/home.vue"
import OneIconsHotel from "../../components/icons/hotel.vue"
import OneIconsRestaurant from "../../components/icons/restaurant.vue"
import OneIconsManagement from "../../components/icons/management.vue"
import OneIconsCircleUserMale from "../../components/icons/circle-user-male.vue"
import OneIconsCircleUserFamale from "../../components/icons/circle-user-famale.vue"

// Errors
import OneIconsError404 from "../../components/icons/error/404.vue"

// Corp
import OneIconsCorpIGS from "../../components/icons/corp/iGS.vue"

function init() {
  // UI Base
  Vue.component('one-icons-menu', OneIconsMenu)
  Vue.component('one-icons-menu-vertical', OneIconsMenuVertical)
  Vue.component('one-icons-circle-user-male', OneIconsCircleUserMale)
  Vue.component('one-icons-circle-user-famale', OneIconsCircleUserFamale)

  // Módulos
  Vue.component('one-icons-home', OneIconsHome)
  Vue.component('one-icons-hotel', OneIconsHotel)
  Vue.component('one-icons-restaurant', OneIconsRestaurant)
  Vue.component('one-icons-management', OneIconsManagement)

  // Errores
  Vue.component('one-icons-error-404', OneIconsError404)

  //Corp
  Vue.component('one-icons-corp-igs', OneIconsCorpIGS)
}

export default init
