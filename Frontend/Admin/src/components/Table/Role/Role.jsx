import { Tooltip, Zoom } from '@mui/material';
import Fade from '@mui/material/Grow';
import Pagination from '@mui/material/Pagination';
import Paper from '@mui/material/Paper';
import Stack from '@mui/material/Stack';
import axios from 'axios';
import { useSnackbar } from 'notistack';
import PropTypes from 'prop-types';
import React, { useContext, useState } from 'react';
import { context } from '../../../app/Context';
import '../stylesTable.scss';
import UpdateRole from './../../UpdateComponent/UpdateRole';
TableRole.propTypes = {
  List: PropTypes.array,
};
TableRole.defaultProps = {
  List: [],
};

export default function TableRole(props) {
  const Context = useContext(context);
  const { List, paginate, setPaginate, setFlag } = props;
  const { enqueueSnackbar } = useSnackbar();
  const { setBodyAdmin } = Context;
  const [open, setOpen] = useState(false);
  const [details, setDetails] = useState({});
  const ListTitleHead = [
    { Name: 'Mã số' },
    { Name: 'Role Name' },
    { Name: 'Mô tả' },
    { Name: 'Xóa' },
    { Name: 'Cập nhật' },
  ];

  function handleDetaits(params) {
    setOpen(true);
    setDetails(params);
  }

  const HandleDelete = async (id) => {
    if (window.confirm('Bạn đã chắc chắn muốn xóa?')) {
      await axios.delete(`/product/delete/${id}`).then(function (response) {
        if (response.status === 200) {
          setFlag(true);
          enqueueSnackbar('Xóa thành công', { variant: 'success' });
        } else {
          enqueueSnackbar('Xóa thất bại', { variant: 'warning' });
        }
      });
    }
  };
  function changePage(page) {
    setFlag(true);
    setPaginate({
      ...paginate,
      page: page,
    });
  }
  function HandelUpdate(id) {
    setBodyAdmin(<UpdateRole id={id} />);
  }
  return (
    <>
      <button
        type='button'
        onClick=''
        className='btn btn-outline-success'
        style={{ position: 'absolute', right: '5%', top: '2%' }}>
        Thêm nhóm quyền
      </button>

      <Stack className='mt-4' spacing={2}>
        <Pagination
          color='primary'
          count={paginate?.count}
          onChange={(e, value) => changePage(value)}
        />
      </Stack>
      <Fade in={true} timeout={400} className='body_page'>
        <Paper>
          <div>
            <table className='itemTable'>
              <thead className='headerTable'>
                <tr>
                  <th>STT</th>
                  {ListTitleHead?.map((item, index) => (
                    <th key={index}>{item?.Name}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {List?.map((item, index) => (
                  <tr key={index} id={item?.Id}>
                    <td>{index + 1}</td>
                    <td>{item?.Id}</td>
                    <td>
                      <p>{item?.RoleName}</p>
                    </td>
                    <td>
                      <Tooltip
                        TransitionComponent={Zoom}
                        title={item?.Description}
                        placement='right-start'
                        arrow>
                        <p className='text_over'>{item?.Description}</p>
                      </Tooltip>
                    </td>

                    <td>
                      <button
                        type='button'
                        className='btn btn-outline-danger'
                        data-set={item?.Id}
                        onClick={() => HandleDelete(item?.Id)}>
                        Xóa
                      </button>
                    </td>
                    <td>
                      <button
                        type='button'
                        className='btn btn-outline-success'
                        onClick={() => HandelUpdate(item?.Id)}
                        data-set={item?.Id}>
                        Cập nhật
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Paper>
      </Fade>
    </>
  );
}
