import React, { memo, useEffect } from "react";
function NotFound() {
  // eslint-disable-next-line react-hooks/exhaustive-deps
  // const click= async()=> axios({
  //     method: 'post',
  //     url: '/news/add',
  //     data: {
  //       Title: 'Thu6',
  //       Content: 'Flintqwqwstone dsadadadda dadadasda dadsadad',
  //       Thumbnail:"https://feed.thecoffeehouse.com/content/images/2021/08/img_8668_grande.jpg"
  //     }
  //   });

  useEffect(() => {
    window.scrollTo(0, 0);
  }, []);
  return (
    <div className="body_Page">
      <img
        src="https://images.all-free-download.com/images/graphiclarge/error_404_page_not_found_6845510.jpg"
        alt=""
        style={{ width: "100%", height: "100%" }}
      />
    </div>
  );
}

export default memo(NotFound);
